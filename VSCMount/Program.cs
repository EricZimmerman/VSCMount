using System;
using System.CommandLine;
using System.CommandLine.Help;
using System.CommandLine.NamingConventionBinder;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Threading.Tasks;
using Exceptionless;
using NLog;
using NLog.Config;
using NLog.Targets;

namespace VSCMount;

internal class Program
{
    private static Logger _loggerConsole;
    private static readonly string BaseDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        
    private static readonly string Header =
        $"VSCMount version {Assembly.GetExecutingAssembly().GetName().Version}" +
        "\r\n\r\nAuthor: Eric Zimmerman (saericzimmerman@gmail.com)" +
        "\r\nhttps://github.com/EricZimmerman/VSCMount";

    private static  string Footer =
        @"Examples: VSCMount.exe --dl C --mp C:\VssRoot --ud --debug" +
        "\r\n\t " +
        "\r\n\t" +
        "    Short options (single letter) are prefixed with a single dash. Long commands are prefixed with two dashes";

    private static RootCommand _rootCommand;

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

    private static async Task Main(string[] args)
    {
        ExceptionlessClient.Default.Startup("vKFCtHS0H467sgdMz3ZqhVoLYF8IZpOCfv1Q38xM");

        SetupNLog();

        _loggerConsole = LogManager.GetLogger("Console");

        if (IsAdministrator() == false)
        {
            _loggerConsole.Fatal("Administrator privileges not found! Exiting!!\r\n");
            return;
        }
            
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            _loggerConsole.Error($"Mounting VSCs only supported on Windows. Exiting\r\n");
            return ;
        }
            
        _rootCommand = new RootCommand
        {
            new Option<string>(
                "--dl",
                "Source drive to look for Volume Shadow Copies (C, D:, or F:\\ for example)"),

            new Option<string>(
                "--mp",
                "The base directory where you want VSCs mapped to"),

            new Option<bool>(
                "--ud",
                () => true,
                "Use VSC creation timestamps (yyyyMMddTHHmmss.fffffff) in symbolic link names"),

            new Option<bool>(
                "--debug",
                () => false,
                "Show debug information during processing"),

        };

        _rootCommand.Description = Header + "\r\n\r\n" + Footer;

        _rootCommand.Handler = CommandHandler.Create(DoWork);

        await _rootCommand.InvokeAsync(args);
    }

    private static void DoWork(string dl, string mp, bool ud, bool debug)
    {
        if (string.IsNullOrEmpty(dl))
        {
            var helpBld = new HelpBuilder(LocalizationResources.Instance, Console.WindowWidth);
            var hc = new HelpContext(helpBld,_rootCommand,Console.Out);

            helpBld.Write(hc);
                
            _loggerConsole.Warn("\r\ndl is required. Exiting\r\n");
            return;
        }

        if (string.IsNullOrEmpty(mp))
        {
            var helpBld = new HelpBuilder(LocalizationResources.Instance, Console.WindowWidth);
            var hc = new HelpContext(helpBld,_rootCommand,Console.Out);

            helpBld.Write(hc);
                
            _loggerConsole.Warn("\r\nmp is required. Exiting\r\n");
            return;
        }

        dl = dl[0].ToString().ToUpperInvariant();

        if (DriveInfo.GetDrives()
                .Any(t => t.RootDirectory.Name.StartsWith(dl)) ==
            false)
        {
            _loggerConsole.Error(
                $"\r\n'{dl}' is not ready. Exiting\r\n");
            return;
        }

        if (debug)
        {
            foreach (var r in LogManager.Configuration.LoggingRules)
            {
                r.EnableLoggingForLevel(LogLevel.Debug);
            }
        }

        LogManager.ReconfigExistingLoggers();

        _loggerConsole.Info(Header);
        _loggerConsole.Info("");

        _loggerConsole.Info($"Command line: {string.Join(" ", Environment.GetCommandLineArgs().Skip(1))}");
        _loggerConsole.Info("");

        if (Directory.Exists(
                $"{mp}_{dl}") ==
            false)
        {
            _loggerConsole.Info(
                $"Creating directory '{mp}_{dl}'");

            try
            {
                Directory.CreateDirectory(
                    $"{mp}_{dl}");
            }
            catch (Exception e)
            {
                _loggerConsole.Fatal($"Unable to create directory '{mp}_{dl}'. Does the drive exist? Error: {e.Message} Exiting\r\n");

                return;
            }
        }

        try
        {
            var vssDirs =
                Directory.GetDirectories(
                    $"{mp}_{dl}",
                    "vss*");

            _loggerConsole.Debug(
                $"Cleaning up vss* directories in '{mp}_{dl}'");
            foreach (var vssDir in vssDirs)
            {
                _loggerConsole.Debug($"Deleting '{vssDir}'");
                Directory.Delete(vssDir);
            }

            _loggerConsole.Info(
                $"Mounting VSCs to '{mp}_{dl}'\r\n");

            Helpers.MountVss(dl.Substring(0, 1),
                $"{mp}_{dl}",
                ud);

            _loggerConsole.Info(
                $"\r\nMounting complete. Navigate VSCs via symbolic links in '{mp}_{dl}'");

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
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            return true;
        }
            
        _loggerConsole.Debug("Checking for admin rights");
        var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }
}