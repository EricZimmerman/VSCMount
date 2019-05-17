using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Principal;
using Exceptionless;
using Fclp;
using NLog;
using NLog.Config;
using NLog.Targets;

namespace VSCMount
{
    internal class Program
    {
        private static Logger _loggerConsole;
        private static FluentCommandLineParser<ApplicationArguments> _fluentCommandLineParser;
        private static readonly string BaseDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        private static void SetupNLog()
        {
            if (File.Exists(Path.Combine(BaseDirectory, "Nlog.config")))
            {
                return;
            }

            var config = new LoggingConfiguration();
            var logLevel = LogLevel.Info;

            const string layout = @"${message}";

            var consoleTarget = new ColoredConsoleTarget();

            config.AddTarget("console", consoleTarget);

            consoleTarget.Layout = layout;

            var rule1 = new LoggingRule("Console", logLevel, consoleTarget);
            config.LoggingRules.Add(rule1);

            LogManager.Configuration = config;
        }

        private static void Main(string[] args)
        {
            ExceptionlessClient.Default.Startup("vKFCtHS0H467sgdMz3ZqhVoLYF8IZpOCfv1Q38xM");

            SetupNLog();

            _loggerConsole = LogManager.GetLogger("Console");

            if (IsAdministrator() == false)
            {
                _loggerConsole.Fatal("Administrator privileges not found! Exiting!!\r\n");
                return;
            }

            _fluentCommandLineParser = new FluentCommandLineParser<ApplicationArguments>
            {
                IsCaseSensitive = false
            };

            _fluentCommandLineParser.Setup(arg => arg.DriveLetter)
                .As("dl")
                .WithDescription("Source drive to look for Volume Shadow Copies (C, D:, or F:\\ for example)");

            _fluentCommandLineParser.Setup(arg => arg.MountPoint)
                .As("mp")
                .WithDescription("The base directory where you want VSCs mapped to");

            _fluentCommandLineParser.Setup(arg => arg.UseDatesInNames)
                .As("ud")
                .WithDescription(
                    "Use VSC creation timestamps (yyyyMMddTHHmmss.fffffff) in symbolic link names. Default is FALSE")
                .SetDefault(false);

            _fluentCommandLineParser.Setup(arg => arg.Debug)
                .As("debug")
                .WithDescription("Show debug information during processing").SetDefault(false);

            var header =
                $"VSCMount version {Assembly.GetExecutingAssembly().GetName().Version}" +
                "\r\n\r\nAuthor: Eric Zimmerman (saericzimmerman@gmail.com)" +
                "\r\nhttps://github.com/EricZimmerman/VSCMount";

            const string footer =
                @"Examples: VSCMount.exe --dl C --mp C:\VssRoot --debug" +
                "\r\n\t " +
                "\r\n\t" +
                "  Short options (single letter) are prefixed with a single dash. Long commands are prefixed with two dashes\r\n";

            _fluentCommandLineParser.SetupHelp("?", "help")
                .WithHeader(header)
                .Callback(text => _loggerConsole.Info(text + "\r\n" + footer));

            var result = _fluentCommandLineParser.Parse(args);

            if (result.HelpCalled)
            {
                return;
            }

            if (result.HasErrors)
            {
                _fluentCommandLineParser.HelpOption.ShowHelp(_fluentCommandLineParser.Options);

                _loggerConsole.Error("");
                _loggerConsole.Error(result.ErrorText);

                return;
            }

            if (string.IsNullOrEmpty(_fluentCommandLineParser.Object.DriveLetter))
            {
                _fluentCommandLineParser.HelpOption.ShowHelp(_fluentCommandLineParser.Options);
                _loggerConsole.Warn("\r\ndl is required. Exiting\r\n");
                return;
            }

            if (string.IsNullOrEmpty(_fluentCommandLineParser.Object.MountPoint))
            {
                _fluentCommandLineParser.HelpOption.ShowHelp(_fluentCommandLineParser.Options);
                _loggerConsole.Warn("\r\nmp is required. Exiting\r\n");
                return;
            }

            _fluentCommandLineParser.Object.DriveLetter =
                _fluentCommandLineParser.Object.DriveLetter[0].ToString().ToUpperInvariant();

            if (DriveInfo.GetDrives()
                    .Any(t => t.RootDirectory.Name.StartsWith(_fluentCommandLineParser.Object.DriveLetter)) ==
                false)
            {
                _loggerConsole.Error(
                    $"\r\n'{_fluentCommandLineParser.Object.DriveLetter}' is not ready. Exiting\r\n");
                return;
            }

            if (_fluentCommandLineParser.Object.Debug)
            {
                foreach (var r in LogManager.Configuration.LoggingRules)
                {
                    r.EnableLoggingForLevel(LogLevel.Debug);
                }
            }

            LogManager.ReconfigExistingLoggers();

            _loggerConsole.Info(header);
            _loggerConsole.Info("");

            _loggerConsole.Info($"Command line: {string.Join(" ", args)}");
            _loggerConsole.Info("");

            if (Directory.Exists(
                    $"{_fluentCommandLineParser.Object.MountPoint}_{_fluentCommandLineParser.Object.DriveLetter}") ==
                false)
            {
                _loggerConsole.Info(
                    $"Creating directory '{_fluentCommandLineParser.Object.MountPoint}_{_fluentCommandLineParser.Object.DriveLetter}'");
                Directory.CreateDirectory(
                    $"{_fluentCommandLineParser.Object.MountPoint}_{_fluentCommandLineParser.Object.DriveLetter}");
            }

            try
            {
                var vssDirs =
                    Directory.GetDirectories(
                        $"{_fluentCommandLineParser.Object.MountPoint}_{_fluentCommandLineParser.Object.DriveLetter}",
                        "vss*");

                _loggerConsole.Debug(
                    $"Cleaning up vss* directories in '{_fluentCommandLineParser.Object.MountPoint}_{_fluentCommandLineParser.Object.DriveLetter}'");
                foreach (var vssDir in vssDirs)
                {
                    _loggerConsole.Debug($"Deleting '{vssDir}'");
                    Directory.Delete(vssDir);
                }

                _loggerConsole.Info(
                    $"Mounting VSCs to '{_fluentCommandLineParser.Object.MountPoint}_{_fluentCommandLineParser.Object.DriveLetter}'\r\n");

                Helpers.MountVss(_fluentCommandLineParser.Object.DriveLetter.Substring(0, 1),
                    $"{_fluentCommandLineParser.Object.MountPoint}_{_fluentCommandLineParser.Object.DriveLetter}",
                    _fluentCommandLineParser.Object.UseDatesInNames);

                _loggerConsole.Info(
                    $"\r\nMounting complete. Navigate VSCs via symbolic links in '{_fluentCommandLineParser.Object.MountPoint}_{_fluentCommandLineParser.Object.DriveLetter}'");

                _loggerConsole.Warn(
                    "\r\nTo remove VSC access, delete individual VSC directories or the main mountpoint directory\r\n");


            }
            catch (Exception ex)
            {
                _loggerConsole.Error(ex,
                    $"Error when mounting VSCs: {ex.Message}");
            }
        }


        private static bool IsAdministrator()
        {
            _loggerConsole.Debug("Checking for admin rights");
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
    }


    internal class ApplicationArguments
    {
        public string DriveLetter { get; set; }
        public string MountPoint { get; set; }
        public bool Debug { get; set; }
        public bool UseDatesInNames { get; set; }
    }
}