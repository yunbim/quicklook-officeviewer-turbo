// Copyright © 2017-2026 QL-Win Contributors
//
// This file is part of QuickLook program.
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.

using QuickLook.Common.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Threading;

namespace QuickLook.Plugin.OfficeViewer;

/// <summary>
/// Which <c>IInitializeWith*</c> flavour a preview handler actually implemented.
/// The Office handlers only implement <see cref="File"/>; probing Stream/Item every
/// single preview is wasted work (it opens the file and creates a Shell item for nothing).
/// </summary>
internal enum InitKind
{
    Unknown = 0,
    File,
    Item,
    Stream,
    None,
}

/// <summary>
/// Owns the lifetime of the out-of-process COM class factories which host the Office
/// preview handlers (EXCEL.EXE / WINWORD.EXE / POWERPNT.EXE).
///
/// <para><b>Why this exists.</b> Measured on a stock Windows + Office install, the three
/// Office preview handlers are registered with a <c>LocalServer32</c> that points straight
/// at the Office executables and implement <b>only</b> <c>IInitializeWithFile</c>
/// (<c>CLSCTX_INPROC_SERVER</c> returns <c>0x80040154</c>). A cold activation therefore
/// costs 480-863 ms because it literally starts Excel/Word/PowerPoint, while a warm
/// activation costs ~2 ms — a 200-400x difference. Caching the <c>IClassFactory</c> is
/// what buys the fast path.</para>
///
/// <para><b>The leak this fixes.</b> The original code called <c>LockServer(true)</c> and
/// never paired it with <c>LockServer(false)</c> anywhere in the repository, so the three
/// Office processes stayed resident forever (~355 MB working set) and kept pushing the
/// file cache out from under the rest of the system. This pool adds the missing half: an
/// idle watchdog unlocks and releases the factories once no preview has run for
/// <see cref="IdleTimeoutSeconds"/> seconds, letting Office shut itself down.</para>
///
/// <para><b>Threading.</b> A COM class factory obtained from <c>CoGetClassObject</c> is
/// bound to the apartment of the thread that obtained it, and QuickLook always drives the
/// preview from its WPF UI thread (STA). Every method here therefore has to run on that
/// same UI thread — including warm-up, which is why warm-up is queued onto the
/// <see cref="Dispatcher"/> at <see cref="DispatcherPriority.Background"/> rather than
/// farmed out to a thread-pool thread. Doing it on a background thread would hand back an
/// interface pointer that is illegal to use from the UI apartment.</para>
/// </summary>
internal static class OfficePreviewHostPool
{
    private const uint ClsctxLocalServer = 0x4;

    private static readonly Guid IidIClassFactory = new("00000001-0000-0000-C000-000000000046");

    private static readonly Guid IidIUnknown = new("00000000-0000-0000-C000-000000000046");

    /// <summary>Default idle timeout before Office hosts are released, in seconds.</summary>
    internal const int DefaultIdleTimeoutSeconds = 300;

    /// <summary>How often the idle watchdog runs, in milliseconds.</summary>
    private const int WatchdogIntervalMs = 30_000;

    /// <summary>Delay between two warm-up activations, in milliseconds.</summary>
    private const int WarmUpSpacingMs = 1200;

    private static readonly object Gate = new();

    private static readonly Dictionary<Guid, Entry> Entries = [];

    private static readonly Dictionary<Guid, InitKind> InitPreferences = [];

    private static readonly HashSet<Guid> WarmUpQueue = [];

    private static Timer _watchdog;

    private static bool _watchdogStarted;

    private static int _idleTimeoutSeconds = DefaultIdleTimeoutSeconds;

    [DllImport("ole32.dll", ExactSpelling = true, PreserveSig = true)]
    private static extern int CoGetClassObject(
        ref Guid rclsid,
        uint dwClsContext,
        IntPtr pvReserved,
        ref Guid riid,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppv);

    private sealed class Entry
    {
        public Guid Clsid;

        /// <summary>The RCW we hold for the class factory (keeps the server alive).</summary>
        public object Rcw;

        public IClassFactory Factory;

        /// <summary>True when the CLSID has no LocalServer registration at all.</summary>
        public bool LocalServerUnavailable;

        public DateTime LastUsedUtc;
    }

    /// <summary>
    /// Applies the user's idle-reclaim configuration. Called once from the plugin.
    /// A value of 0 disables reclamation (Office stays resident, as before).
    /// </summary>
    internal static void Configure(int idleTimeoutSeconds)
    {
        lock (Gate)
        {
            _idleTimeoutSeconds = idleTimeoutSeconds;
            EnsureWatchdogLocked();
        }
    }

    /// <summary>
    /// Returns a cached class factory for <paramref name="clsid"/>, creating (and thereby
    /// cold-starting) the server if this is the first use. Returns null when the CLSID has
    /// no out-of-process registration, in which case the caller may try in-process.
    /// Must be called on the UI thread.
    /// </summary>
    internal static IClassFactory TryGetFactory(Guid clsid, out bool localServerUnavailable)
    {
        lock (Gate)
        {
            EnsureWatchdogLocked();

            if (Entries.TryGetValue(clsid, out var cached))
            {
                cached.LastUsedUtc = DateTime.UtcNow;
                localServerUnavailable = cached.LocalServerUnavailable;
                return cached.Factory;
            }

            var entry = new Entry { Clsid = clsid, LastUsedUtc = DateTime.UtcNow };

            var factoryIid = IidIClassFactory;
            var hr = CoGetClassObject(ref clsid, ClsctxLocalServer, IntPtr.Zero, ref factoryIid, out var rcw);
            if (hr < 0 || rcw == null)
            {
                entry.LocalServerUnavailable = true;
                Entries[clsid] = entry;
                localServerUnavailable = true;
                return null;
            }

            entry.Rcw = rcw;
            entry.Factory = (IClassFactory)rcw;

            // Lock the server so it survives between previews. This is paired by
            // ReleaseEntry() — the pairing the original code was missing.
            try
            {
                entry.Factory.LockServer(true);
            }
            catch (COMException e)
            {
                ProcessHelper.WriteLog($"OfficePreviewHostPool: LockServer failed for {clsid:B}: {e.Message}");
            }

            Entries[clsid] = entry;
            localServerUnavailable = false;
            return entry.Factory;
        }
    }

    /// <summary>
    /// Creates a handler instance from the cached factory, evicting a dead host and retrying
    /// once. Sets <paramref name="localServerUnavailable"/> when the CLSID turned out to have
    /// no out-of-process registration at all. Must be called on the UI thread.
    /// </summary>
    internal static object CreateHandlerInstance(Guid clsid, out bool localServerUnavailable)
    {
        var factory = TryGetFactory(clsid, out localServerUnavailable);
        if (factory == null)
            return null;

        for (var attempt = 0; attempt < 2; attempt++)
        {
            try
            {
                var iid = IidIUnknown;
                factory.CreateInstance(null, ref iid, out var instance);
                return instance;
            }
            catch (COMException)
            {
                // The host process probably died between previews. Drop the stale entry so
                // the next attempt cold-starts a fresh one.
                Evict(clsid);
                factory = TryGetFactory(clsid, out localServerUnavailable);
                if (factory == null)
                    return null;
            }
        }

        return null;
    }

    /// <summary>
    /// Evicts a single CLSID, releasing the factory reference so a dead server can be
    /// replaced.
    /// </summary>
    private static void Evict(Guid clsid)
    {
        Entry stale = null;

        lock (Gate)
        {
            if (!Entries.TryGetValue(clsid, out var entry))
                return;

            stale = new Entry { Clsid = clsid, Rcw = entry.Rcw, Factory = entry.Factory };
            Entries.Remove(clsid);
        }

        Release(stale);
    }

    /// <summary>
    /// The init-interface a handler for <paramref name="clsid"/> turned out to support, so
    /// later previews can skip the flavours that are known to fail.
    /// </summary>
    internal static InitKind GetInitPreference(Guid clsid)
    {
        lock (Gate)
        {
            return InitPreferences.TryGetValue(clsid, out var kind) ? kind : InitKind.Unknown;
        }
    }

    internal static void SetInitPreference(Guid clsid, InitKind kind)
    {
        lock (Gate)
        {
            InitPreferences[clsid] = kind;
        }
    }

    /// <summary>
    /// Schedules a background-priority warm-up of the given Office handler CLSIDs.
    ///
    /// The heavy part of a cold preview is <c>CoGetClassObject</c> starting the Office
    /// executable. Doing that at <see cref="DispatcherPriority.Background"/> keeps it out
    /// of the user-visible path: the dispatcher only gets here once the window has been
    /// shown and nothing else is queued, so by the time the user picks the next file the
    /// server is already up and activation collapses to ~2 ms.
    ///
    /// Requests are coalesced and already-warm CLSIDs are skipped, so calling this
    /// liberally is cheap.
    /// </summary>
    internal static void RequestWarmUp(IEnumerable<Guid> clsids)
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher == null || dispatcher.HasShutdownStarted)
            return;

        var added = false;
        lock (Gate)
        {
            foreach (var clsid in clsids)
            {
                if (clsid == Guid.Empty)
                    continue;

                // Already warm, or already queued.
                if (Entries.TryGetValue(clsid, out var entry) && entry.Factory != null)
                    continue;

                if (WarmUpQueue.Add(clsid))
                    added = true;
            }
        }

        if (!added)
            return;

        dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(RunWarmUp));
    }

    private static void RunWarmUp()
    {
        Guid next;
        lock (Gate)
        {
            if (WarmUpQueue.Count == 0)
                return;

            using var it = WarmUpQueue.GetEnumerator();
            it.MoveNext();
            next = it.Current;
            WarmUpQueue.Remove(next);
        }

        try
        {
            TryGetFactory(next, out _);
        }
        catch (Exception e)
        {
            ProcessHelper.WriteLog($"OfficePreviewHostPool: warm-up of {next:B} failed: {e.Message}");
        }

        // Yield the UI thread between activations so a warm-up never turns into one long
        // freeze, then come back at Background priority for the next one.
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher != null && !dispatcher.HasShutdownStarted && IsWarmUpPending())
        {
            var timer = new DispatcherTimer(DispatcherPriority.Background, dispatcher)
            {
                Interval = TimeSpan.FromMilliseconds(WarmUpSpacingMs),
            };
            timer.Tick += (s, _) =>
            {
                ((DispatcherTimer)s).Stop();
                RunWarmUp();
            };
            timer.Start();
        }
    }

    private static bool IsWarmUpPending()
    {
        lock (Gate)
        {
            return WarmUpQueue.Count > 0;
        }
    }

    private static void EnsureWatchdogLocked()
    {
        if (_watchdogStarted)
            return;

        _watchdogStarted = true;
        _watchdog = new Timer(OnWatchdogTick, null, WatchdogIntervalMs, WatchdogIntervalMs);
    }

    /// <summary>
    /// Watchdog entry point. Runs on a thread-pool thread, so it does no COM work itself —
    /// it just hops onto the UI thread, which is the apartment that owns the factories.
    /// </summary>
    private static void OnWatchdogTick(object state)
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher == null || dispatcher.HasShutdownStarted)
            return;

        dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(ReclaimIdleEntries));
    }

    /// <summary>
    /// Unlocks and releases every factory that has been idle for longer than the configured
    /// timeout. Must run on the UI thread (see the threading note on the class).
    /// </summary>
    private static void ReclaimIdleEntries()
    {
        var timeout = _idleTimeoutSeconds;
        if (timeout <= 0)
            return;

        List<Entry> stale = null;
        var now = DateTime.UtcNow;

        lock (Gate)
        {
            foreach (var kv in Entries)
            {
                var entry = kv.Value;
                if (entry.Factory == null || (now - entry.LastUsedUtc).TotalSeconds < timeout)
                    continue;

                // Take ownership of the references and clear them from the entry in the
                // same locked step, so nothing can hand them out again. The references
                // themselves have to be captured here — releasing them must happen outside
                // the lock, and by then the live entry no longer points at them.
                (stale ??= []).Add(new Entry
                {
                    Clsid = entry.Clsid,
                    Rcw = entry.Rcw,
                    Factory = entry.Factory,
                });
                entry.Rcw = null;
                entry.Factory = null;
            }
        }

        if (stale == null)
            return;

        foreach (var entry in stale)
            Release(entry);

        ProcessHelper.WriteLog(
            $"OfficePreviewHostPool: released {stale.Count} idle Office preview host(s) after {timeout}s " +
            $"idle: {string.Join(", ", stale.Select(s => s.Clsid.ToString("B")))}");
    }

    /// <summary>
    /// Unlocks the server and drops our class-factory reference. This is the pair for
    /// <c>LockServer(true)</c>; with it the Office process is free to shut down.
    /// </summary>
    private static void Release(Entry entry)
    {
        if (entry == null)
            return;

        if (entry.Factory != null)
        {
            try
            {
                entry.Factory.LockServer(false);
            }
            catch (Exception e)
            {
                ProcessHelper.WriteLog($"OfficePreviewHostPool: LockServer(false) failed: {e.Message}");
            }
        }

        if (entry.Rcw != null)
        {
            try
            {
                Marshal.FinalReleaseComObject(entry.Rcw);
            }
            catch (Exception)
            {
                // ignored
            }
        }

        entry.Rcw = null;
        entry.Factory = null;
    }
}
