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

using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace QuickLook.Plugin.OfficeViewer;

/// <summary>
/// A Windows Forms host for Shell preview handlers.
///
/// <para>Activation is delegated to <see cref="OfficePreviewHostPool"/>, which caches the
/// out-of-process class factory and reclaims it when idle.</para>
///
/// <para><b>Note on the initialization order.</b> This used to always probe
/// <c>IInitializeWithStream</c> → <c>IInitializeWithItem</c> → <c>IInitializeWithFile</c>.
/// Forensics on a stock install show the three Office handlers implement <b>only</b>
/// <c>IInitializeWithFile</c>, so the first two probes could never succeed: every preview
/// paid for a <c>File.OpenRead</c> and a <c>SHCreateItemFromParsingName</c> that were
/// guaranteed to be thrown away. The flavour that actually worked is now remembered per
/// CLSID, so steady-state previews go straight to the one that matters — and non-Office
/// handlers (e.g. the Visio surrogate) keep the original behaviour via the fallback order.</para>
/// </summary>
public class PreviewHandlerHost : Control
{
    private static readonly Guid IidIShellItem = new("43826d1e-e718-42ee-bc55-a1e261c37bfe");

    private IPreviewHandler _mCurrentPreviewHandler;
    private Stream _fileStream;

    /// <summary>
    /// Initialize a new instance of the PreviewHandlerHost class.
    /// </summary>
    public PreviewHandlerHost()
    {
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, PreserveSig = true)]
    private static extern int SHCreateItemFromParsingName(
        [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
        IntPtr pbc,
        ref Guid riid,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppv);

    /// <summary>
    /// Gets the GUID of the current preview handler.
    /// </summary>
    [Browsable(false)]
    [ReadOnly(true)]
    public Guid CurrentPreviewHandler { get; private set; } = Guid.Empty;

    /// <summary>
    /// Releases the unmanaged resources used by the PreviewHandlerHost and optionally releases the managed resources.
    /// </summary>
    protected override void Dispose(bool disposing)
    {
        UnloadPreviewHandler();

        _fileStream?.Dispose();
        _fileStream = null;

        if (_mCurrentPreviewHandler != null)
        {
            Marshal.FinalReleaseComObject(_mCurrentPreviewHandler);
            _mCurrentPreviewHandler = null;

            // NOTE: a blocking GC.Collect() used to live here. This path runs on the WPF UI
            // thread every time a preview window closes, so it stalled the UI for the
            // duration of a full GC. Releasing the RCW is sufficient; the runtime reclaims
            // the managed wrapper on its own schedule.
        }

        base.Dispose(disposing);
    }

    /// <summary>
    /// Resizes the hosted preview handler when this PreviewHandlerHost is resized.
    /// </summary>
    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);

        try
        {
            var r = ClientRectangle;
            _mCurrentPreviewHandler?.SetRect(ref r);
        }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x8001010D))
        {
            // RPC_E_CANTCALLOUT_ININPUTSYNCCALL
            // This exception occurs when an outgoing call cannot be made because
            // the application is dispatching an input-synchronous call.
            // It's safe to ignore this exception as the preview handler will be
            // resized on the next resize event.
        }
    }

    /// <summary>
    /// Opens the specified file using the appropriate preview handler and displays the result
    /// in this PreviewHandlerHost. The initialization flavour that works for the handler is
    /// tried first and remembered for subsequent previews.
    /// </summary>
    public bool Open(string path)
    {
        UnloadPreviewHandler();
        _fileStream?.Dispose();
        _fileStream = null;

        if (string.IsNullOrEmpty(path))
            return false;

        var guid = ShellExRegister.GetPreviewHandlerGUID(Path.GetExtension(path));
        if (guid == Guid.Empty)
            return false;

        CurrentPreviewHandler = guid;

        // Out-of-process first: the Office handlers are registered as LocalServer32 against
        // the Office executables, and the pool keeps that server warm between previews.
        var o = OfficePreviewHostPool.CreateHandlerInstance(guid, out var localServerUnavailable);

        // In-process is kept purely as a fallback for CLSIDs with no LocalServer
        // registration. It is dead code for Office itself — CLSCTX_INPROC_SERVER on the
        // three Office handlers returns 0x80040154 (class not registered).
        if (o == null && localServerUnavailable)
            o = TryCreateInProcHandler(guid);

        if (o == null)
            return false;

        if (!TryInitializeHandler(guid, o, path))
        {
            Marshal.FinalReleaseComObject(o);
            return false;
        }

        _mCurrentPreviewHandler = o as IPreviewHandler;
        if (_mCurrentPreviewHandler == null)
        {
            Marshal.FinalReleaseComObject(o);
            return false;
        }

        if (IsDisposed)
            return false;

        var r = ClientRectangle;
        _mCurrentPreviewHandler.SetWindow(Handle, ref r);
        _mCurrentPreviewHandler.DoPreview();

        return true;
    }

    /// <summary>
    /// Unloads the preview handler hosted in this PreviewHandlerHost.
    /// </summary>
    public void UnloadPreviewHandler()
    {
        try
        {
            _mCurrentPreviewHandler?.Unload();
        }
        catch (Exception)
        {
            // ignored
        }
    }

    /// <summary>
    /// Initializes <paramref name="handler"/>, trying the flavour remembered for this CLSID
    /// first and only then falling back to the remaining ones. The flavour that succeeds is
    /// cached so later previews skip the probes that cannot work.
    /// </summary>
    private bool TryInitializeHandler(Guid clsid, object handler, string path)
    {
        var preferred = OfficePreviewHostPool.GetInitPreference(clsid);

        var order = preferred switch
        {
            InitKind.File => new[] { InitKind.File, InitKind.Item, InitKind.Stream },
            InitKind.Item => new[] { InitKind.Item, InitKind.Stream, InitKind.File },
            InitKind.Stream => new[] { InitKind.Stream, InitKind.Item, InitKind.File },
            _ => new[] { InitKind.Stream, InitKind.Item, InitKind.File },
        };

        foreach (var kind in order)
        {
            if (TryInitializeWith(kind, handler, path))
            {
                if (preferred != kind)
                    OfficePreviewHostPool.SetInitPreference(clsid, kind);

                return true;
            }
        }

        return false;
    }

    private bool TryInitializeWith(InitKind kind, object handler, string path)
    {
        const uint stgmRead = 0;

        try
        {
            switch (kind)
            {
                case InitKind.Stream:
                    if (handler is not IInitializeWithStream streamInit)
                        return false;

                    var stream = File.OpenRead(path);
                    try
                    {
                        streamInit.Initialize(new IStreamWrapper(stream), stgmRead);
                    }
                    catch
                    {
                        stream.Dispose();
                        throw;
                    }

                    _fileStream = stream;
                    return true;

                case InitKind.Item:
                    if (handler is not IInitializeWithItem itemInit)
                        return false;

                    var riid = IidIShellItem;
                    var hr = SHCreateItemFromParsingName(path, IntPtr.Zero, ref riid, out var shellItemObj);
                    if (hr < 0)
                        return false;

                    itemInit.Initialize((IShellItem)shellItemObj, stgmRead);
                    return true;

                case InitKind.File:
                    if (handler is not IInitializeWithFile fileInit)
                        return false;

                    fileInit.Initialize(path, stgmRead);
                    return true;

                default:
                    return false;
            }
        }
        catch (Exception)
        {
            // Probe failed — the caller moves on to the next flavour.
            return false;
        }
    }

    /// <summary>
    /// Falls back to in-process COM instantiation via <see cref="Activator.CreateInstance"/>.
    /// </summary>
    private static object TryCreateInProcHandler(Guid clsid)
    {
        try
        {
            return Activator.CreateInstance(Type.GetTypeFromCLSID(clsid, true));
        }
        catch
        {
            return null;
        }
    }
}
