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
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using UnblockZoneIdentifier;

namespace QuickLook.Plugin.OfficeViewer;

/// <summary>
/// Preview support for documents carrying a <c>Zone.Identifier</c> mark ("downloaded from the
/// Internet"), which Windows surfaces as Protected View.
///
/// <para>The Office preview handlers refuse to load a marked file, so the mark has to go
/// before the document can be rendered at all. QuickLook used to remove that mark from the
/// <b>original</b> file — i.e. previewing a document silently rewrote it, and afterwards Word
/// would no longer open it in Protected View either. That is surprising for what is supposed
/// to be a read-only peek. Instead a temporary copy is streamed into
/// <c>%TEMP%\QuickLook\OfficeViewer</c> and only that copy is unblocked: the original is
/// opened for reading and never modified.</para>
///
/// <para>The copy is written byte-by-byte into a brand-new file, so it cannot inherit the
/// source's alternate data stream in the first place. The unblocking below is therefore a
/// verification step, not the mechanism that makes the preview work.</para>
/// </summary>
internal static class ProtectedViewPreview
{
    /// <summary>Copies in the cache are deleted once they are older than this.</summary>
    public static readonly TimeSpan CopyRetention = TimeSpan.FromDays(1);

    private static readonly string CacheDirectory =
        Path.Combine(Path.GetTempPath(), "QuickLook", "OfficeViewer");

    /// <summary>Name of the NTFS alternate data stream holding the zone mark.</summary>
    private const string ZoneStream = ":Zone.Identifier";

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DeleteFileW(string lpFileName);

    /// <summary>
    /// True when the file carries an Internet zone mark and would open in Protected View.
    /// </summary>
    public static bool IsProtected(string path) => IsBlocked(path);

    /// <summary>
    /// Produces an unblocked copy of <paramref name="sourcePath"/> and returns its path, or
    /// <c>null</c> when the copy could not be created. The source is opened read-only with
    /// maximally permissive sharing so that previewing never disturbs whoever else has it open.
    /// </summary>
    public static string CreateUnblockedCopy(string sourcePath)
    {
        try
        {
            Directory.CreateDirectory(CacheDirectory);

            var target = Path.Combine(CacheDirectory, BuildCopyName(sourcePath));

            try
            {
                WriteCopy(sourcePath, target);
            }
            catch (IOException)
            {
                // A warm Office host may still hold the previous copy of this very file open.
                // Fall back to a unique name rather than failing the preview.
                target = Path.Combine(CacheDirectory, BuildCopyName(sourcePath, unique: true));
                WriteCopy(sourcePath, target);
            }

            // The copy is clean by construction; confirm it, because a copy that is *still*
            // blocked would present to the user as exactly the bug this code exists to avoid.
            if (IsBlocked(target))
                RemoveMark(target);

            return target;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Removes the Internet zone mark from <paramref name="path"/> in place. Only ever called
    /// on an explicit user request (the banner button), never automatically.
    /// </summary>
    public static bool UnblockSource(string path)
    {
        try
        {
            if (!IsBlocked(path))
                return true;

            RemoveMark(path);
            return !IsBlocked(path);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>Best-effort removal of a copy that is no longer needed.</summary>
    public static void TryDelete(string path)
    {
        if (string.IsNullOrEmpty(path))
            return;

        try
        {
            File.Delete(path);
        }
        catch
        {
            // Still mapped by an Office host; the age based sweep picks it up later.
        }
    }

    /// <summary>
    /// Deletes cached copies left behind by earlier sessions. Called once, when the plugin is
    /// loaded, so it never sits on the path of a preview.
    /// </summary>
    public static void SweepStaleCopies()
    {
        try
        {
            if (!Directory.Exists(CacheDirectory))
                return;

            var cutoff = DateTime.UtcNow - CopyRetention;

            foreach (var file in Directory.GetFiles(CacheDirectory))
            {
                try
                {
                    if (File.GetLastWriteTimeUtc(file) < cutoff)
                        File.Delete(file);
                }
                catch
                {
                    // In use by a warm Office host; it will be swept on a later run.
                }
            }
        }
        catch
        {
            // Housekeeping must never break startup.
        }
    }

    private static void WriteCopy(string source, string target)
    {
        using var input = new FileStream(
            source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var output = new FileStream(
            target, FileMode.Create, FileAccess.Write, FileShare.Read);

        input.CopyTo(output, 81920);
    }

    /// <summary>
    /// Deterministic file name for the copy, derived from the source path plus its size and
    /// timestamp so that editing the document produces a new name instead of a stale hit.
    /// </summary>
    private static string BuildCopyName(string sourcePath, bool unique = false)
    {
        var info = new FileInfo(sourcePath);
        var stamp = string.Concat(
            sourcePath, "|",
            info.Length.ToString(), "|",
            info.LastWriteTimeUtc.Ticks.ToString());

        string tag;
        using (var sha = SHA256.Create())
        {
            var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(stamp));
            var hex = new StringBuilder(12);
            for (var i = 0; i < 6; i++)
                hex.Append(hash[i].ToString("x2"));
            tag = hex.ToString();
        }

        if (unique)
            tag += "-" + Guid.NewGuid().ToString("N").Substring(0, 6);

        var name = Path.GetFileNameWithoutExtension(sourcePath);
        foreach (var invalid in Path.GetInvalidFileNameChars())
            name = name.Replace(invalid, '_');
        if (name.Length > 60)
            name = name.Substring(0, 60);

        return string.Concat(name, "-", tag, Path.GetExtension(sourcePath));
    }

    private static bool IsBlocked(string path)
    {
        try
        {
            return ZoneIdentifierManager.IsZoneBlocked(path);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Clears the zone mark. The packaged API is tried first so the behaviour stays consistent
    /// with how QuickLook detects the mark; deleting the alternate data stream directly is the
    /// deterministic fallback (it is exactly what "Unblock" in the file properties does).
    /// </summary>
    private static void RemoveMark(string path)
    {
        try
        {
            ZoneIdentifierManager.RemoveZone(path);
        }
        catch
        {
            // Fall through to the direct implementation.
        }

        if (!IsBlocked(path))
            return;

        try
        {
            DeleteFileW(path + ZoneStream);
        }
        catch
        {
            // Reported to the caller via the verification in UnblockSource.
        }
    }
}
