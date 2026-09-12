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
using System.Globalization;
using QuickLook.Common.Helpers;

namespace QuickLook.Plugin.OfficeViewer;

/// <summary>
/// Bilingual (Chinese / English) text for everything this patch draws itself.
///
/// <para>Everything QuickLook already drew — the window frame, the other plugins — is untouched;
/// only the strings added by this patch live here.</para>
///
/// <para>The language is taken from the Windows UI language of the current user
/// (<see cref="CultureInfo.CurrentUICulture"/>), so a Chinese Windows gets Chinese and every other
/// system gets English with no configuration. It can be forced with the <c>UiLanguage</c> setting
/// (<c>auto</c> / <c>zh</c> / <c>en</c>) — see the README.</para>
/// </summary>
internal static class UiText
{
    private const string SettingDomain = "QuickLook.Plugin.OfficeViewer";

    private static bool? _chinese;

    private static bool Chinese
    {
        get
        {
            if (_chinese == null)
            {
                _chinese = Resolve();
            }

            return _chinese.Value;
        }
    }

    private static bool Resolve()
    {
        // Explicit setting wins.
        try
        {
            var preference = SettingHelper.Get("UiLanguage", "auto", SettingDomain);
            if (!string.IsNullOrWhiteSpace(preference))
            {
                switch (preference.Trim().ToLowerInvariant())
                {
                    case "zh":
                    case "cn":
                    case "chs":
                    case "zh-cn":
                    case "zh-hans":
                        return true;
                    case "en":
                    case "eng":
                    case "en-us":
                        return false;
                }
            }
        }
        catch
        {
            // A missing or unreadable config file must never break a preview.
        }

        // Otherwise follow the Windows UI language.
        try
        {
            return CultureInfo.CurrentUICulture.TwoLetterISOLanguageName
                .Equals("zh", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>Picks one of two strings by the resolved UI language.</summary>
    public static string T(string chinese, string english)
    {
        return Chinese ? chinese : english;
    }

    // ---- Protected View notice strip ---------------------------------------

    public static string BannerMessage => T(
        "副本预览 — 该文件来自 Internet，Office 无法在“受保护的视图”下加载，"
        + "所以这里显示的是临时副本，内容与原文件一致。你的原文件没有被修改。",
        "Copy preview — this file came from the Internet, so Office refuses to load it "
        + "in Protected View. What you see is a temporary copy of it. Your original file "
        + "has not been modified.");

    public static string BannerActionLabel => T(
        "解除源文件保护视图",
        "Unblock the original file");

    public static string BannerActionTip => T(
        "删除该文件上的 Zone.Identifier 标记，等同于在文件属性里点“解除锁定”。只影响这一个文件。",
        "Removes the Zone.Identifier mark from that one file — the same as ticking "
        + "\"Unblock\" in its file properties. Affects only that file.");

    public static string BannerDoneMessage => T(
        "已解除源文件的保护视图标记 — 以后打开它不会再进入受保护的视图。"
        + "文档内容本身没有做任何改动。",
        "Protected View mark removed from the original file — it will now open normally. "
        + "The document content itself was not touched.");

    public static string BannerDoneTag => T("✓ 已解除", "✓ Unblocked");

    public static string BannerFailureMessage => T(
        "解除失败：文件可能被占用或权限不足。"
        + "也可以右键该文件 → 属性 → 勾选“解除锁定”。",
        "Could not unblock: the file may be in use or you may lack permission. You can "
        + "also right-click the file → Properties → tick \"Unblock\".");

    public static string BannerRetryLabel => T("重试", "Retry");

    // ---- Viewer title -------------------------------------------------------

    public static string TitleCopyPreview(string fileName)
    {
        return T($"[副本预览] {fileName}", $"[Copy preview] {fileName}");
    }

    public static string TitleProtectedView(string fileName)
    {
        return $"[PROTECTED VIEW] {fileName}";
    }

    public static string NoCopyAvailable => T(
        "该文档受 Office 保护，且无法创建用于预览的临时副本。",
        "This document is protected by Office and no temporary copy could be created "
        + "for previewing it.");
}
