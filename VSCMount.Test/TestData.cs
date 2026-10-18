using System;
using System.IO;
using System.Security.Principal;
using NUnit.Framework;

namespace VSCMount.Test;

internal static class TestData
{
    public const string ShadowCopyTarget = @"\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1\";

    public static string NewTempDir()
    {
        var dir = Path.Combine(Path.GetTempPath(), "VSCMount.Test", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    public static bool IsWindowsAdministrator =>
        OperatingSystem.IsWindows() &&
        new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);

    /// <summary>
    ///     Creates a directory symbolic link (which needs administrator rights or Developer Mode on Windows), else ignores
    ///     the test
    /// </summary>
    public static void CreateLinkOrIgnore(string path, string target)
    {
        try
        {
            Directory.CreateSymbolicLink(path, target);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            Assert.Ignore($"Can't create symbolic links here ({e.Message})");
        }
    }

    /// <summary>
    ///     Removes links first, so a recursive delete can never reach what they point to
    /// </summary>
    public static void DeleteDir(string dir)
    {
        if (!Directory.Exists(dir)) return;

        foreach (var entry in new DirectoryInfo(dir).EnumerateFileSystemInfos("*", SearchOption.TopDirectoryOnly))
        {
            if (entry.Attributes.HasFlag(FileAttributes.ReparsePoint))
                entry.Delete();
            else if (entry is DirectoryInfo subDir)
                DeleteDir(subDir.FullName);
            else
                entry.Delete();
        }

        Directory.Delete(dir);
    }
}
