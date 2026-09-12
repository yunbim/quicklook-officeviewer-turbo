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
using QuickLook.Common.Plugin;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using UnblockZoneIdentifier;

namespace QuickLook.Plugin.OfficeViewer;

public sealed class Plugin : IViewer
{
    private const string SettingDomain = "QuickLook.Plugin.OfficeViewer";

    private static readonly string[] Extensions =
    [
        ".doc", ".docx", ".docm", ".odt",
        ".xls", ".xlsx", ".xlsm", ".xlsb", ".ods",
        ".ppt", ".pptx", ".odp",
        ".vsd", ".vsdx",
    ];

    /// <summary>
    /// Representative extensions used to discover which preview handlers exist on this
    /// machine, so the warm-up can pre-activate exactly those and skip the rest.
    /// </summary>
    private static readonly string[] WarmUpExtensions = [".xlsx", ".docx", ".pptx", ".vsdx"];

    private PreviewPanel _panel;

    /// <summary>
    /// Path of the temporary, unblocked copy backing the current preview, or <c>null</c> when
    /// the document was previewed in place. Kept only so it can be cleaned up.
    /// </summary>
    private string _temporaryCopy;

    public int Priority => -1;

    public void Init()
    {
        // Applies the idle-reclaim policy and (optionally) pre-activates the Office preview
        // hosts. Init runs once, on the UI thread, while the plugin manager is loading.
        OfficePreviewHostPool.Configure(
            SettingHelper.Get("HandlerIdleTimeoutSeconds",
                OfficePreviewHostPool.DefaultIdleTimeoutSeconds, SettingDomain));

        if (SettingHelper.Get("WarmUpAtStartup", false, SettingDomain))
            OfficePreviewHostPool.RequestWarmUp(OfficeHandlerClsids());

        // Housekeeping for previews of Protected View documents, which are served from a
        // temporary copy. Done here rather than per preview so it can never sit on the
        // critical path of the first preview.
        ProtectedViewPreview.SweepStaleCopies();
    }

    /// <summary>
    /// The preview handler CLSIDs actually registered for this machine's Office formats.
    /// </summary>
    private static IEnumerable<Guid> OfficeHandlerClsids()
    {
        return WarmUpExtensions
            .Select(ShellExRegister.GetPreviewHandlerGUID)
            .Where(guid => guid != Guid.Empty)
            .Distinct();
    }

    public bool CanHandle(string path)
    {
        if (Directory.Exists(path))
            return false;

        if (!Extensions.Any(path.ToLower().EndsWith))
            return false;

        var previewHandler = ShellExRegister.GetPreviewHandlerGUID(Path.GetExtension(path));
        if (previewHandler == Guid.Empty)
            return false;

        var checkPreviewHandler = SettingHelper.Get("CheckPreviewHandler", true, "QuickLook.Plugin.OfficeViewer");
        if (!checkPreviewHandler)
            return true;

        if (!string.IsNullOrWhiteSpace(CLSIDRegister.GetName(previewHandler.ToString("B"))))
        {
            return true;
        }
        else
        {
            // Legacy: No more recovering registries for MS Office.
            // TODO: Add a setting page to let users choose the fallback preview handler if the current one is not working.
#if false
            // To restore the preview handler CLSID to MS Office
            // if running with administrative privileges
            if (ShellExRegister.IsRunAsAdmin())
            {
                var fileExtension = Path.GetExtension(path);
                var fallbackHandler = fileExtension switch
                {
                    ".doc" or ".docx" or ".docm" or ".odt" => CLSIDRegister.MicrosoftWord,
                    ".xls" or ".xlsx" or ".xlsm" or ".xlsb" or ".ods" => CLSIDRegister.MicrosoftExcel,
                    ".ppt" or ".pptx" or ".odp" => CLSIDRegister.MicrosoftPowerPoint,
                    ".vsd" or ".vsdx" => CLSIDRegister.MicrosoftVisio,
                    _ => null,
                };

                if (fallbackHandler == null)
                    return false;

                if (!string.IsNullOrWhiteSpace(CLSIDRegister.GetName(fallbackHandler)))
                {
                    // Admin requested
                    ShellExRegister.SetPreviewHandlerGUID(fileExtension, new Guid(fallbackHandler));
                    return true;
                }
            }
#endif
        }

        return false;
    }

    public void Prepare(string path, ContextObject context)
    {
        context.SetPreferredSizeFit(new Size { Width = 1200, Height = 800 }, 0.8d);
    }

    public void View(string path, ContextObject context)
    {
        // MS Office's preview handlers refuse to load a document that carries an Internet zone
        // mark, so such a file is previewed through an unblocked temporary copy instead of
        // having its own mark removed. Previewing used to rewrite the original document; now it
        // is only ever read, and the user decides from the notice whether to unblock it for real.
        var previewPath = path;
        string temporaryCopy = null;

        if (ProtectedViewPreview.IsProtected(path))
        {
            temporaryCopy = ProtectedViewPreview.CreateUnblockedCopy(path);

            if (temporaryCopy == null)
            {
                // Copying failed (no temp access, disk full, ...). Say so, rather than handing
                // the handler a file it is bound to refuse and showing a blank pane.
                context.ViewerContent = new Label()
                {
                    Content = UiText.NoCopyAvailable,
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                };
                context.Title = UiText.TitleProtectedView(Path.GetFileName(path));
                context.IsBusy = false;
                return;
            }

            previewPath = temporaryCopy;
        }

        try
        {
            _temporaryCopy = temporaryCopy;
            _panel = new PreviewPanel();

            if (temporaryCopy != null)
            {
                // The notice needs its own layout row above the preview: a WindowsFormsHost
                // paints over any WPF sibling it overlaps, so it cannot be floated on top.
                context.ViewerContent = new ProtectedViewBanner(path, _panel);
                context.Title = UiText.TitleCopyPreview(Path.GetFileName(path));
            }
            else
            {
                context.ViewerContent = _panel;
                context.Title = Path.GetFileName(path);
            }

            var stopwatch = Stopwatch.StartNew();
            var opened = _panel.PreviewFile(previewPath, context);
            stopwatch.Stop();

            ReportPreviewTiming(Path.GetExtension(path), stopwatch.ElapsedMilliseconds);

            // The first Office preview is unavoidably cold: it starts the Office executable
            // (measured 482-863 ms). Kick off the remaining handlers now so that browsing a
            // folder of mixed documents never pays that cost a second time.
            if (opened && SettingHelper.Get("WarmUpAfterFirstPreview", true, SettingDomain))
                OfficePreviewHostPool.RequestWarmUp(OfficeHandlerClsids());
        }
        catch (Exception e)
        {
            context.ViewerContent = new Label()
            {
                Content = e.ToString(),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
            };
        }

        context.IsBusy = false;
    }

    /// <summary>
    /// Emits how long the synchronous part of a preview took.
    ///
    /// This is the bit that was missing before: Release builds had no timing data anywhere
    /// in this pipeline, which is why "preview is slow" could not be attributed to the COM
    /// activation step. Always goes to the debugger output; set <c>LogPreviewTiming</c> to
    /// also append it to the exception log for post-hoc analysis.
    /// </summary>
    private static void ReportPreviewTiming(string extension, long elapsedMs)
    {
        var line = $"OfficeViewer: {extension} preview (handler activation + DoPreview) took {elapsedMs} ms";
        Debug.WriteLine(line);

        if (SettingHelper.Get("LogPreviewTiming", false, SettingDomain))
            ProcessHelper.WriteLog(line);
    }

    public void Cleanup()
    {
        _panel?.Dispose();
        _panel = null;

        // Best effort: the copy is often still mapped by the warm Office host at this point, in
        // which case the delete fails silently and the age based sweep in Init clears it later.
        ProtectedViewPreview.TryDelete(_temporaryCopy);
        _temporaryCopy = null;
    }
}
