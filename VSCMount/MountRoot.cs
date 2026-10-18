using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Serilog;

namespace VSCMount;

public static class MountRoot
{
    /// <summary>
    ///     Creates <paramref name="root" />, or removes the <c>vss*</c> links an earlier run left in it. When it holds
    ///     anything else, nothing is deleted and false is returned with the reason in <paramref name="error" />, so only
    ///     links VSCMount created are ever removed
    /// </summary>
    public static bool Prepare(string root, out string error)
    {
        error = null;

        if (File.Exists(root))
        {
            error = $"{root} is a file";
            return false;
        }

        if (!Directory.Exists(root))
        {
            Log.Information("Creating directory {Root}", root);
            Directory.CreateDirectory(root);
            return true;
        }

        var links = new List<FileSystemInfo>();
        var others = new List<string>();

        foreach (var entry in new DirectoryInfo(root).EnumerateFileSystemInfos())
        {
            if (IsShadowCopyLink(entry))
                links.Add(entry);
            else
                others.Add(entry.Name);
        }

        if (others.Count > 0)
        {
            var examples = string.Join(", ", others.Take(5));
            if (others.Count > 5) examples += ", ...";

            error =
                $"{root} contains {others.Count:N0} item(s) VSCMount did not create ({examples}). Nothing was deleted. Use a new or empty --mp folder";
            return false;
        }

        foreach (var link in links)
        {
            //Deletes the link only, never what it points to
            Log.Debug("Removing link {Link} from an earlier run", link.FullName);
            link.Delete();
        }

        return true;
    }

    /// <summary>
    ///     True for a <c>vss*</c> symbolic link to a shadow copy device, the only thing VSCMount creates
    /// </summary>
    public static bool IsShadowCopyLink(FileSystemInfo entry)
    {
        return entry.Name.StartsWith("vss", StringComparison.OrdinalIgnoreCase) &&
               entry.Attributes.HasFlag(FileAttributes.ReparsePoint) &&
               VssParsing.IsShadowCopyTarget(entry.LinkTarget);
    }
}
