using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using Serilog;
using WmiLight;

namespace VSCMount;

public class VssInfo
{
    public VssInfo(int vssNumber, DateTimeOffset? createdOn, string shadowCopyId, string shadowCopyVolume,
        string originatingMachine, string servicingMachine, char volumeLetter, string originalVolume)
    {
        VssNumber = vssNumber;
        CreatedOn = createdOn;
        ShadowCopyId = shadowCopyId?.ToLowerInvariant();
        ShadowCopyVolume = shadowCopyVolume;
        OriginatingMachine = originatingMachine;
        ServicingMachine = servicingMachine;
        VolumeLetter = volumeLetter;
        OriginalVolume = originalVolume;
    }

    /// <summary>UTC. Null when WMI's value can't be read</summary>
    public DateTimeOffset? CreatedOn { get; }

    public string ShadowCopyId { get; }
    public string ShadowCopyVolume { get; }
    public string OriginatingMachine { get; }
    public string ServicingMachine { get; }
    public char VolumeLetter { get; }
    public string OriginalVolume { get; }
    public int VssNumber { get; }
}

public static class Helpers
{
    private const uint SymbolicLinkFlagDirectory = 1;

    /// <summary>
    ///     Win32_Volume DeviceID to Caption (C:\ for a volume with a drive letter). Needs no administrator rights
    /// </summary>
    internal static Dictionary<string, string> GetVolumes()
    {
        using var con = new WmiConnection();
        return GetVolumes(con);
    }

    private static Dictionary<string, string> GetVolumes(WmiConnection con)
    {
        var volumes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var vol in con.CreateQuery("SELECT Caption,DeviceID FROM Win32_Volume"))
        {
            var deviceId = vol["DeviceID"]?.ToString();
            var caption = vol["Caption"]?.ToString();

            if (deviceId != null && caption != null) volumes[deviceId] = caption;
        }

        return volumes;
    }

    //CreateSymbolicLinkW returns BOOLEAN (1 byte), not BOOL, so the return value must be marshaled as I1
    [DllImport("kernel32.dll", EntryPoint = "CreateSymbolicLinkW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.I1)]
    private static extern bool CreateSymbolicLink(string lpSymlinkFileName, string lpTargetFileName, uint dwFlags);

    /// <summary>
    ///     The shadow copies of the volume with <paramref name="driveLetter" />, by number. Copies that can't be used are
    ///     skipped with a warning
    /// </summary>
    public static List<VssInfo> GetVssInfoViaWmi(char driveLetter)
    {
        Log.Debug("Running WMI queries to get VSC info");

        var vss = new List<VssInfo>();

        using var con = new WmiConnection();

        var volumes = GetVolumes(con);

        Log.Verbose("Volume info from WMI: {@Volumes}", volumes);

        foreach (var scInfo in con.CreateQuery(
                     "SELECT DeviceObject,ID,InstallDate,OriginatingMachine,VolumeName,ServiceMachine FROM Win32_ShadowCopy"))
        {
            var id = scInfo["ID"]?.ToString();
            var devObj = scInfo["DeviceObject"]?.ToString();
            var origVolume = scInfo["VolumeName"]?.ToString();

            if (origVolume == null || !volumes.TryGetValue(origVolume, out var caption))
            {
                Log.Warning("VSC {Id} is of volume {Volume}, which WMI does not list. Skipping", id, origVolume);
                continue;
            }

            //The caption is C:\ for a volume with a drive letter, else its mount folder or volume GUID path
            if (VssParsing.NormalizeDriveLetter(caption) != driveLetter) continue;

            var vssNumber = VssParsing.GetShadowCopyNumber(devObj);
            if (vssNumber == null)
            {
                Log.Warning("VSC {Id} has an unexpected device name ({DeviceObject}). Skipping", id, devObj);
                continue;
            }

            var installDate = scInfo["InstallDate"];
            var createdOn = installDate switch
            {
                DateTime dateTime => new DateTimeOffset(dateTime).ToUniversalTime(),
                _ => VssParsing.ParseCimDateTime(installDate?.ToString())
            };

            if (createdOn == null)
                Log.Warning("VSC {Id} has a creation date that can't be read ({InstallDate})", id, installDate);

            var vsI = new VssInfo(vssNumber.Value, createdOn, id, devObj, scInfo["OriginatingMachine"]?.ToString(),
                scInfo["ServiceMachine"]?.ToString(), driveLetter, origVolume);

            Log.Verbose("Adding VSC: {@Vsc}", vsI);
            vss.Add(vsI);
        }

        vss.Sort((a, b) => a.VssNumber.CompareTo(b.VssNumber));

        Log.Debug("Found {Count:N0} VSCs", vss.Count);

        return vss;
    }

    /// <summary>
    ///     Creates a symbolic link in <paramref name="mountRoot" /> for each shadow copy
    /// </summary>
    /// <returns>The number of links that could not be created</returns>
    public static int MountVss(IReadOnlyList<VssInfo> shadowCopies, string mountRoot, bool useDatesInNames)
    {
        var failed = 0;

        foreach (var vssInfo in shadowCopies)
        {
            Log.Debug(
                "Attempting to mount VSS with Id: {ShadowCopyId}, Creation date: {CreatedOn:yyyy-MM-dd HH:mm:ss.fffffff}",
                vssInfo.ShadowCopyId, vssInfo.CreatedOn);

            var mountDir = Path.Combine(mountRoot,
                VssParsing.GetLinkName(vssInfo.VssNumber, vssInfo.CreatedOn, useDatesInNames));

            if (CreateSymbolicLink(mountDir, $@"{vssInfo.ShadowCopyVolume}\", SymbolicLinkFlagDirectory))
            {
                Log.Information(
                    "\tVSS {VssNumber} (Id {ShadowCopyId}, Created on: {CreatedOn:yyyy-MM-dd HH:mm:ss.fffffff} UTC) mounted OK!",
                    vssInfo.VssNumber.ToString().PadRight(4), vssInfo.ShadowCopyId, vssInfo.CreatedOn);
            }
            else
            {
                var reason = new Win32Exception(Marshal.GetLastPInvokeError()).Message;

                Log.Warning(
                    "\tVSS {VssNumber} (Id {ShadowCopyId}, Created on: {CreatedOn:yyyy-MM-dd HH:mm:ss.fffffff} UTC) failed to mount: {Reason}",
                    vssInfo.VssNumber.ToString().PadRight(4), vssInfo.ShadowCopyId, vssInfo.CreatedOn, reason);
                failed++;
            }
        }

        return failed;
    }
}
