using System;
using System.CommandLine;
using System.CommandLine.Help;
using System.CommandLine.Invocation;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Versioning;
using System.Security.Principal;
using System.Threading.Tasks;
using Exceptionless;
using Serilog;
using Serilog.Core;
using Serilog.Events;

namespace VSCMount;

internal class Program
{
    private static readonly string Header =
        $"VSCMount version {Assembly.GetExecutingAssembly().GetName().Version.ToString(3)}" +
        "\r\n\r\nAuthor: Eric Zimmerman (saericzimmerman@gmail.com)" +
        "\r\nhttps://github.com/EricZimmerman/VSCMount";

    private static readonly string Footer =
        @"Examples: VSCMount.exe --dl C --mp C:\VssRoot --debug" +
        "\r\n\t " +
        "\r\n\t" +
        "    Links are created in <mp>_<drive letter> (C:\\VssRoot_C above). Use --ud false for names without dates" +
        "\r\n\t" +
        "    Short options (single letter) are prefixed with a single dash. Long commands are prefixed with two dashes";

    private static RootCommand _rootCommand;

    private static readonly LoggingLevelSwitch _levelSwitch = new();

    private static async Task<int> Main(string[] args)
    {
        ExceptionlessClient.Default.Startup("vKFCtHS0H467sgdMz3ZqhVoLYF8IZpOCfv1Q38xM");

        var template = "{Message:lj}{NewLine}{Exception}";

        var conf = new LoggerConfiguration()
            .WriteTo.Console(outputTemplate: template)
            .MinimumLevel.ControlledBy(_levelSwitch);

        Log.Logger = conf.CreateLogger();

        var dlOpt = new Option<string>("--dl")
        {
            Description = "Source drive to look for Volume Shadow Copies (C, D:, or F:\\ for example)"
        };

        var mpOpt = new Option<string>("--mp")
        {
            Description = "The base directory where you want VSCs mapped to"
        };

        var udOpt = new Option<bool>("--ud")
        {
            Description =
                "Use VSC creation timestamps (yyyyMMddTHHmmss.fffffff, UTC) in symbolic link names. Default is true. Use --ud false to turn off",
            DefaultValueFactory = _ => true
        };

        var debugOpt = new Option<bool>("--debug")
        {
            Description = "Show debug information during processing",
            DefaultValueFactory = _ => false
        };

        var traceOpt = new Option<bool>("--trace")
        {
            Description = "Show trace information during processing",
            DefaultValueFactory = _ => false
        };

        _rootCommand = new RootCommand
        {
            dlOpt,
            mpOpt,
            udOpt,
            debugOpt,
            traceOpt
        };

        _rootCommand.Description = Header + "\r\n\r\n" + Footer;

        _rootCommand.SetAction(result => DoWork(result.GetValue(dlOpt), result.GetValue(mpOpt), result.GetValue(udOpt),
            result.GetValue(debugOpt), result.GetValue(traceOpt)));

        var exitCode = await _rootCommand.Parse(args).InvokeAsync();

        Log.CloseAndFlush();

        return exitCode;
    }

    private class CustomHelpAction : SynchronousCommandLineAction
    {
        private readonly HelpAction _defaultHelp;

        public CustomHelpAction(HelpAction action)
        {
            _defaultHelp = action;
        }

        public override int Invoke(ParseResult parseResult)
        {
            return _defaultHelp.Invoke(parseResult);
        }
    }

    /// <returns>
    ///     0 on success (including a volume without VSCs). 1 for invalid arguments, a drive that isn't ready, missing
    ///     administrator rights, a mount folder that can't be used, a WMI error or a VSC that failed to mount
    /// </returns>
    private static int DoWork(string dl, string mp, bool ud, bool debug, bool trace)
    {
        var template = "{Message:lj}{NewLine}{Exception}";
        if (debug)
        {
            _levelSwitch.MinimumLevel = LogEventLevel.Debug;

            template = "[{Timestamp:HH:mm:ss.fff} {Level:u3}] {Message:lj}{NewLine}{Exception}";
        }

        if (trace)
        {
            _levelSwitch.MinimumLevel = LogEventLevel.Verbose;
            template = "[{Timestamp:HH:mm:ss.fff} {Level:u3}] {Message:lj}{NewLine}{Exception}";
        }

        var conf = new LoggerConfiguration()
            .WriteTo.Console(outputTemplate: template)
            .MinimumLevel.ControlledBy(_levelSwitch);

        Log.Logger = conf.CreateLogger();

        if (string.IsNullOrEmpty(dl) || string.IsNullOrEmpty(mp))
        {
            var helpAction = new CustomHelpAction(new HelpAction());
            helpAction.Invoke(_rootCommand.Parse(""));

            Log.Warning("Both --dl and --mp are required. Exiting");
            Console.WriteLine();
            return 1;
        }

        var driveLetter = VssParsing.NormalizeDriveLetter(dl);
        if (driveLetter == null)
        {
            Log.Error("{Dl} is not a drive letter. Use C, C: or C:\\ for example. Exiting", dl);
            Console.WriteLine();
            return 1;
        }

        Log.Information("{Header}", Header);
        Console.WriteLine();

        Log.Information("Command line: {Args}", string.Join(" ", Environment.GetCommandLineArgs().Skip(1)));
        Console.WriteLine();

        if (!OperatingSystem.IsWindows())
        {
            Log.Fatal("Mounting VSCs is only supported on Windows. Exiting");
            Console.WriteLine();
            return 1;
        }

        if (IsAdministrator() == false)
        {
            Log.Fatal("Administrator privileges not found! Exiting!!");
            Console.WriteLine();
            return 1;
        }

        if (DriveInfo.GetDrives().Any(t => t.Name[0] == driveLetter && t.IsReady) == false)
        {
            Log.Error("{Dl}: is not ready. Exiting", driveLetter);
            Console.WriteLine();
            return 1;
        }

        string mountRoot;
        try
        {
            mountRoot = VssParsing.GetMountRoot(Path.GetFullPath(mp), driveLetter.Value);

            if (MountRoot.Prepare(mountRoot, out var error) == false)
            {
                Log.Fatal("Unable to use {MountRoot}: {Error}. Exiting", mountRoot, error);
                Console.WriteLine();
                return 1;
            }
        }
        catch (Exception e)
        {
            Log.Fatal(e, "Unable to prepare directory for {Mp}. Does the drive exist? Error: {Message} Exiting", mp,
                e.Message);
            Console.WriteLine();
            return 1;
        }

        try
        {
            var shadowCopies = Helpers.GetVssInfoViaWmi(driveLetter.Value);

            if (shadowCopies.Count == 0)
            {
                Log.Warning("No VSCs found on volume {DriveLetter}:. Nothing to mount", driveLetter);
                Console.WriteLine();
                return 0;
            }

            Log.Warning("VSCs found on volume {DriveLetter}: {Count:N0}. Mounting to {MountRoot}...", driveLetter,
                shadowCopies.Count, mountRoot);
            Console.WriteLine();

            var failed = Helpers.MountVss(shadowCopies, mountRoot, ud);

            Console.WriteLine();
            if (failed > 0)
                Log.Error("{Failed:N0} of {Count:N0} VSCs failed to mount. The others can be navigated via symbolic links in {MountRoot}",
                    failed, shadowCopies.Count, mountRoot);
            else
                Log.Information("Mounting complete. Navigate VSCs via symbolic links in {MountRoot}", mountRoot);

            Console.WriteLine();
            Log.Warning("To remove VSC access, delete individual VSC directories or the main mountpoint directory");
            Console.WriteLine();

            return failed > 0 ? 1 : 0;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Error when mounting VSCs: {Message}", ex.Message);
            Console.WriteLine();
            return 1;
        }
    }

    [SupportedOSPlatform("windows")]
    private static bool IsAdministrator()
    {
        Log.Debug("Checking for admin rights");
        var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }
}
