using System;
using System.CommandLine;
using System.CommandLine.Help;
using System.CommandLine.Invocation;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Threading.Tasks;
using Exceptionless;
using Serilog;
using Serilog.Core;
using Serilog.Events;

namespace VSCMount;

internal class Program
{
    private static readonly string BaseDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

    private static readonly string Header =
        $"VSCMount version {Assembly.GetExecutingAssembly().GetName().Version.ToString(3)}" +
        "\r\n\r\nAuthor: Eric Zimmerman (saericzimmerman@gmail.com)" +
        "\r\nhttps://github.com/EricZimmerman/VSCMount";

    private static readonly string Footer =
        @"Examples: VSCMount.exe --dl C --mp C:\VssRoot --ud --debug" +
        "\r\n\t " +
        "\r\n\t" +
        "    Short options (single letter) are prefixed with a single dash. Long commands are prefixed with two dashes";

    private static RootCommand _rootCommand;

    private static readonly LoggingLevelSwitch _levelSwitch = new();

    private static async Task Main(string[] args)
    {
        ExceptionlessClient.Default.Startup("vKFCtHS0H467sgdMz3ZqhVoLYF8IZpOCfv1Q38xM");

        var template = "{Message:lj}{NewLine}{Exception}";

        var conf = new LoggerConfiguration()
            .WriteTo.Console(outputTemplate: template)
            .MinimumLevel.ControlledBy(_levelSwitch);

        Log.Logger = conf.CreateLogger();

        if (IsAdministrator() == false)
        {
            Log.Fatal("Administrator privileges not found! Exiting!!");
            Console.WriteLine();
            return;
        }

        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            Log.Error("Mounting VSCs only supported on Windows. Exiting");
            Console.WriteLine();
            return;
        }

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
            Description = "Use VSC creation timestamps (yyyyMMddTHHmmss.fffffff) in symbolic link names",
            DefaultValueFactory = _ => true
        };
        
        var debugOpt = new Option<bool>("--debug")
        {
            Description = "Show debug information during processing",
            DefaultValueFactory = _ => false
        };
        
        _rootCommand = new RootCommand
        {
           dlOpt,
           mpOpt,
           udOpt,
           debugOpt
        };

        _rootCommand.Description = Header + "\r\n\r\n" + Footer;

        //_rootCommand.Handler = CommandHandler.Create(DoWork);
        _rootCommand.SetAction(result => DoWork(result.GetValue(dlOpt),result.GetValue(mpOpt),result.GetValue(udOpt),result.GetValue(debugOpt)));

        var foo = _rootCommand.Parse(args).InvokeAsync();

        Log.CloseAndFlush();
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
            var result = _defaultHelp.Invoke(parseResult);

            Log.Warning("{Msg}", string.Join(' ',parseResult.Tokens));

            return result;
        }
    }
    
    
   

    private static void DoWork(string dl, string mp, bool ud, bool debug)
    {
        var template = "{Message:lj}{NewLine}{Exception}";
        if (debug)
        {
            _levelSwitch.MinimumLevel = LogEventLevel.Debug;

            template = "[{Timestamp:HH:mm:ss.fff} {Level:u3}] {Message:lj}{NewLine}{Exception}";
        }

        var conf = new LoggerConfiguration()
            .WriteTo.Console(outputTemplate: template)
            .MinimumLevel.ControlledBy(_levelSwitch);

        Log.Logger = conf.CreateLogger();

        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            Console.WriteLine();
            Log.Fatal(
                "Non-Windows platforms not supported due to the need to use specific Windows libraries! Exiting...");
            Console.WriteLine();
            Environment.Exit(0);
            return;
        }

        if (string.IsNullOrEmpty(dl))
        {
            var aaa = new CustomHelpAction(new HelpAction());
            aaa.Invoke(_rootCommand.Parse("--dl is required. Exiting"));

            return;
        }

        if (string.IsNullOrEmpty(mp))
        {
              var aaa = new CustomHelpAction(new HelpAction());
                aaa.Invoke(_rootCommand.Parse("--mp is required. Exiting"));

                return;
        }

        dl = dl[0].ToString().ToUpperInvariant();

        if (DriveInfo.GetDrives()
                .Any(t => t.RootDirectory.Name.StartsWith(dl)) ==
            false)
        {
            Console.WriteLine();
            Log.Error("{Dl} is not ready. Exiting", dl);
            Console.WriteLine();
            return;
        }


        Log.Information("{Header}", Header);
        Console.WriteLine();

        Log.Information("Command line: {Args}", string.Join(" ", Environment.GetCommandLineArgs().Skip(1)));
        Console.WriteLine();

        if (Directory.Exists(
                $"{mp}_{dl}") ==
            false)
        {
            Log.Information("Creating directory {Mp}_{Dl}", mp, dl);

            try
            {
                Directory.CreateDirectory($"{mp}_{dl}");
            }
            catch (Exception e)
            {
                Log.Fatal(e, "Unable to create directory {Mp}_{Dl}. Does the drive exist? Error: {Message} Exiting", mp,
                    dl, e.Message);
                Console.WriteLine();

                return;
            }
        }

        try
        {
            var vssDirs =
                Directory.GetDirectories(
                    $"{mp}_{dl}",
                    "vss*");

            Log.Debug("Cleaning up vss* directories in {Mp}_{Dl}", mp, dl);
            foreach (var vssDir in vssDirs)
            {
                Log.Debug("Deleting {VssDir}", vssDir);
                Directory.Delete(vssDir);
            }

            Log.Information("Mounting VSCs to {Mp}_{Dl}", mp, dl);
            Console.WriteLine();

            Helpers.MountVss(dl.Substring(0, 1),
                $"{mp}_{dl}",
                ud);

            Console.WriteLine();
            Log.Information("Mounting complete. Navigate VSCs via symbolic links in {Mp}_{Dl}", mp, dl);

            Console.WriteLine();
            Log.Warning("To remove VSC access, delete individual VSC directories or the main mountpoint directory");
            Console.WriteLine();
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Error when mounting VSCs: {Message}", ex.Message);
        }
    }

    private static bool IsAdministrator()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return true;

        Log.Debug("Checking for admin rights");
        var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }
}