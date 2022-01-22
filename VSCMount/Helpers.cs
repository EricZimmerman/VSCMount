using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Serilog;
using ServiceStack.Text;
using VSCMount;
using WmiLight;

namespace VSCMount
{
    public class VssInfo
    {
        public VssInfo(DateTimeOffset createdOn, string shadowCopyId, string shadowCopyVolume,
            string originatingMachine, string servicingMachine, string volumeLetter, string originalVolume)
        {
            CreatedOn = createdOn;
            ShadowCopyId = shadowCopyId.ToLowerInvariant();
            ShadowCopyVolume = shadowCopyVolume;
            OriginatingMachine = originatingMachine;
            ServicingMachine = servicingMachine;
            VolumeLetter = volumeLetter;
            OriginalVolume = originalVolume;

            VssNumber = int.Parse(
                shadowCopyVolume.Substring(shadowCopyVolume.IndexOf("VolumeShadowCopy", StringComparison.Ordinal) +
                                           16));
        }

        public DateTimeOffset CreatedOn { get; }
        public string ShadowCopyId { get; }
        public string ShadowCopyVolume { get; }
        public string OriginatingMachine { get; }
        public string ServicingMachine { get; }
        public string VolumeLetter { get; }
        public string OriginalVolume { get; }
        public int VssNumber { get; }
    }
}

public class Helpers
{
    public enum SymbolicLink
    {
        File = 0,
        Directory = 1
    }

    [DllImport("kernel32.dll")]
    public static extern bool CreateSymbolicLink(
        string lpSymlinkFileName, string lpTargetFileName, SymbolicLink dwFlags);

    public static List<VssInfo> GetVssInfoViaWmi(string driveLetter)
    {
        if (driveLetter == null) driveLetter = string.Empty;

        if (driveLetter.Length > 1) driveLetter = driveLetter.Substring(0, 1);


        var vss = new List<VssInfo>();

        Log.Debug("Running WMI queries to get VSC info");

        var volInfo = new Dictionary<string, string>();

        using (var con = new WmiConnection())
        {
            foreach (var vol in con.CreateQuery("SELECT caption,DeviceID FROM Win32_volume"))
                volInfo.Add(vol["DeviceID"].ToString(), vol["caption"].ToString());
        }

        Log.Verbose("Volume info from WMI: {VolInfo}", volInfo.Dump());

        using (var con = new WmiConnection())
        {
            foreach (var scInfo in con.CreateQuery(
                         "SELECT DeviceObject,ID,InstallDate,OriginatingMachine,VolumeName,ServiceMachine FROM Win32_ShadowCopy")
                    )
            {
                var devObj = scInfo["DeviceObject"].ToString();
                var id = scInfo["ID"].ToString();
                var installDate = scInfo["InstallDate"].ToString();

                var instDateTimeOffset = DateTimeOffset.ParseExact(installDate.Substring(0, installDate.Length - 4),
                    "yyyyMMddHHmmss.ffffff", null, DateTimeStyles.AssumeLocal).ToUniversalTime();

                var origMachine = scInfo["OriginatingMachine"].ToString();
                var serviceMachine = scInfo["ServiceMachine"].ToString();
                var origVolume = scInfo["VolumeName"].ToString();

                var volLetter = volInfo[origVolume].Substring(0, 1);

                var vsI = new VssInfo(instDateTimeOffset, id, devObj, origMachine, serviceMachine, volLetter,
                    origVolume);

                if (!volLetter.ToUpperInvariant().StartsWith(driveLetter.ToUpperInvariant()) &&
                    driveLetter.Trim().Length != 0)
                    continue;

                Log.Verbose("Adding VSC: {Vsc}", vsI.Dump());
                vss.Add(vsI);
            }
        }

        Log.Debug("Found {Count:N0} VSCs", vss.Count);

        return vss;
    }

    public static void MountVss(string driveLetter, string mountRoot, bool useDatesInNames)
    {
        var existingVss = GetVssInfoViaWmi(driveLetter);

        Log.Warning(
            "VSCs found on volume {DriveLetter}: {Count:N0}. Mounting...", driveLetter.ToUpperInvariant(),
            existingVss.Count);

        if (Directory.Exists(mountRoot))
        {
            Log.Debug("mountRoot directory exists. Deleting...");
            foreach (var directory in Directory.GetDirectories(mountRoot)) Directory.Delete(directory, true);

            Directory.Delete(mountRoot, true);
        }

        if (Directory.Exists(mountRoot) == false)
        {
            Log.Debug("Creating mountRoot directory: {MountRoot}", mountRoot);
            Directory.CreateDirectory(mountRoot);
        }

        foreach (var vssInfo in existingVss)
        {
            Log.Debug(
                "Attempting to mount VSS with Id: {ShadowCopyId}, Creation date: {CreatedOn:yyyy-MM-dd HH:mm:ss.fffffff}",
                vssInfo.ShadowCopyId, vssInfo.CreatedOn);

            var mountDir = $@"{mountRoot}\vss{vssInfo.VssNumber:000}";

            if (useDatesInNames) mountDir = $"{mountDir}-{vssInfo.CreatedOn:yyyyMMddTHHmmss.fffffff}";

            var worked = CreateSymbolicLink(mountDir, $@"{vssInfo.ShadowCopyVolume}\", SymbolicLink.Directory);

            if (worked)
                Log.Information(
                    "\tVSS {VssNumber} (Id {ShadowCopyId}, Created on: {CreatedOn:yyyy-MM-dd HH:mm:ss.fffffff} UTC) mounted OK!",
                    vssInfo.VssNumber.ToString().PadRight(4), vssInfo.ShadowCopyId, vssInfo.CreatedOn);
            else
                Log.Warning(
                    "\tVSS {VssNumber} (Id {ShadowCopyId}, Created on: {CreatedOn:yyyy-MM-dd HH:mm:ss.fffffff} UTC) failed to mount!",
                    vssInfo.VssNumber.ToString().PadRight(4), vssInfo.ShadowCopyId, vssInfo.CreatedOn);
        }
    }
}