// Creates and reads Windows .lnk shortcuts.
//
// Why this exists: creating a shortcut normally means instantiating the WScript.Shell COM
// object, and COM instantiation is blocked in this environment's shell tools. QuickLook
// already ships a managed IShellLinkW/IPersistFile wrapper (repo/QuickLook/NativeMethods/
// ShellLink.cs), so this small tool compiles that wrapper in and drives it directly - the
// same code path QuickLook itself uses for its "start with Windows" shortcut.

using System;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using QuickLook.NativeMethods;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.Error.WriteLine("usage: ShortcutMaker create <target.exe> <out.lnk> [arguments]");
            Console.Error.WriteLine("       ShortcutMaker read <in.lnk>");
            return 2;
        }

        try
        {
            switch (args[0])
            {
                case "create" when args.Length >= 3:
                    return Create(args[1], args[2], args.Length > 3 ? args[3] : null);

                case "read" when args.Length >= 2:
                    return Read(args[1]);

                default:
                    Console.Error.WriteLine("bad arguments");
                    return 2;
            }
        }
        catch (Exception e)
        {
            Console.Error.WriteLine(e.Message);
            return 1;
        }
    }

    private static int Create(string target, string lnkPath, string arguments)
    {
        if (!File.Exists(target))
        {
            Console.Error.WriteLine("target does not exist: " + target);
            return 1;
        }

        var lnk = (IShellLinkW)new ShellLink();
        lnk.SetPath(target);
        if (!string.IsNullOrEmpty(arguments))
            lnk.SetArguments(arguments);
        lnk.SetIconLocation(target, 0);
        lnk.SetWorkingDirectory(Path.GetDirectoryName(target));
        lnk.SetDescription("QuickLook - press Space to preview a file");

        var directory = Path.GetDirectoryName(lnkPath);
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);

        ((IPersistFile)lnk).Save(lnkPath, false);

        Console.WriteLine("created: " + lnkPath);
        return 0;
    }

    private static int Read(string lnkPath)
    {
        if (!File.Exists(lnkPath))
        {
            Console.Error.WriteLine("no such shortcut: " + lnkPath);
            return 1;
        }

        var lnk = (IShellLinkW)new ShellLink();
        ((IPersistFile)lnk).Load(lnkPath, 0);

        var path = new StringBuilder(1024);
        lnk.GetPath(path, path.Capacity, out _, SLGP_FLAGS.SLGP_RAWPATH);

        var working = new StringBuilder(1024);
        lnk.GetWorkingDirectory(working, working.Capacity);

        var icon = new StringBuilder(1024);
        lnk.GetIconLocation(icon, icon.Capacity, out var iconIndex);

        Console.WriteLine("target      : " + path);
        Console.WriteLine("working dir : " + working);
        Console.WriteLine("icon        : " + icon + "," + iconIndex);
        return 0;
    }
}
