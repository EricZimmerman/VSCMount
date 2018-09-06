using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using NLog;

namespace VSCMount
{
    public class VssInfo
    {
        public VssInfo(DateTimeOffset createdOn, string shadowCopyId, string shadowCopyVolume,
            string originatingMachine, string servicingMachine)
        {
            CreatedOn = createdOn;
            ShadowCopyId = shadowCopyId;
            ShadowCopyVolume = shadowCopyVolume;
            OriginatingMachine = originatingMachine;
            ServicingMachine = servicingMachine;

            VssNumber = int.Parse(
                shadowCopyVolume.Substring(shadowCopyVolume.IndexOf("VolumeShadowCopy", StringComparison.Ordinal) +
                                           16));
        }

        public DateTimeOffset CreatedOn { get; }
        public string ShadowCopyId { get; }
        public string ShadowCopyVolume { get; }
        public string OriginatingMachine { get; }
        public string ServicingMachine { get; }
        public int VssNumber { get; }

        public override string ToString()
        {
            return
                $"Vss#: {VssNumber}, Created on: {CreatedOn:yyyy/MM/dd HH:mm:ss}, Id: {ShadowCopyId}, Volume: {ShadowCopyVolume}, Origin machine: {OriginatingMachine}, Servicing machine: {ServicingMachine}";
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

        public static List<VssInfo> GetVssForVolume(string driveLetter)
        {
            var loggerConsole = LogManager.GetLogger("Console");

            var vss = new List<VssInfo>();

            using (var p = new Process())
            {
                p.StartInfo.FileName = "vssadmin.exe";
                p.StartInfo.Arguments = $"list shadows /for={driveLetter}:";

                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardError = true;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.CreateNoWindow = true;

                p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

                p.Start();

                var stdOut = p.StandardOutput.ReadToEnd();

                p.WaitForExit();

                var regexObj1 = new Regex("shadow copies at (creation time: .+?)Provider",
                    RegexOptions.Singleline | RegexOptions.Compiled);
                var matchResults1 = regexObj1.Match(stdOut);
                while (matchResults1.Success)
                {
                    for (var i = 1; i <= matchResults1.Groups.Count; i++)
                    {
                        var groupObj = matchResults1.Groups[i];
                        if (!groupObj.Success)
                        {
                            continue;
                        }

                        var createTimeRaw = Regex.Match(groupObj.Value, "creation time: (.+)").Groups[1].Value
                            .TrimEnd();
                        var shadowCopyId = Regex.Match(groupObj.Value, "Shadow Copy ID: (.+)").Groups[1].Value
                            .TrimEnd();
                        var shadowCopyVolume = Regex.Match(groupObj.Value, "Shadow Copy Volume: (.+)").Groups[1].Value
                            .TrimEnd();
                        var originatingMachine = Regex.Match(groupObj.Value, "Originating Machine: (.+)").Groups[1]
                            .Value.TrimEnd();
                        var serviceMachine = Regex.Match(groupObj.Value, "Service Machine: (.+)").Groups[1].Value
                            .TrimEnd();

                        var vi = new VssInfo(
                            DateTimeOffset.Parse(createTimeRaw, null, DateTimeStyles.AdjustToUniversal), shadowCopyId,
                            shadowCopyVolume, originatingMachine, serviceMachine);

                        loggerConsole.Debug($"Adding VSC: {vi}");

                        vss.Add(vi);
                    }

                    matchResults1 = matchResults1.NextMatch();
                }
            }

            loggerConsole.Debug($"Discovered {vss.Count:N0} VSCs");

            return vss;
        }

        public static void MountVss(char driveLetter, string mountRoot, bool useDatesInNames)
        {
            var loggerConsole = LogManager.GetLogger("Console");

            var existingVss = GetVssForVolume(driveLetter.ToString());

            loggerConsole.Warn($"VSCs found on volume {driveLetter}: {existingVss.Count:N0}. Mounting...");

            if (Directory.Exists(mountRoot) == false)
            {
                loggerConsole.Debug($"Creating mountRoot directory: {mountRoot}");
                Directory.CreateDirectory(mountRoot);
            }

            foreach (var vssInfo in existingVss)
            {
                loggerConsole.Debug(
                    $"Attempting to mount VSS with id: {vssInfo.ShadowCopyId}, Creation date: {vssInfo.CreatedOn:yyyy/MM/dd HH:mm:ss}");

                var mountDir = $@"{mountRoot}\vss{vssInfo.VssNumber:000}";

                if (useDatesInNames)
                {
                    mountDir = $"{mountDir}-{vssInfo.CreatedOn:yyyyMMddTHHmmss}";
                }

                var worked = CreateSymbolicLink(mountDir, $@"{vssInfo.ShadowCopyVolume}\", SymbolicLink.Directory);

                if (worked)
                {
                    loggerConsole.Info(
                        $"\tVSS {vssInfo.VssNumber.ToString().PadRight(4)} (Id {vssInfo.ShadowCopyId}, Created on: {vssInfo.CreatedOn:yyyy/MM/dd HH:mm:ss} UTC) mounted OK!");
                }
                else
                {
                    loggerConsole.Warn(
                        $"\tVSS {vssInfo.VssNumber.ToString().PadRight(4)} (Id {vssInfo.ShadowCopyId}, Created on: {vssInfo.CreatedOn:yyyy/MM/dd HH:mm:ss} UTC) failed to mount!");
                }
            }
        }
    }
}