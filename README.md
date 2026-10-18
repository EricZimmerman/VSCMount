# VSCMount

Mounts every Volume Shadow Copy (VSC) of a drive as a folder, so you can browse it with Explorer or point other tools at it.

## Command Line Interface

    Description:
      VSCMount version 2026.10.0

      Author: Eric Zimmerman (saericzimmerman@gmail.com)
      https://github.com/EricZimmerman/VSCMount

      Examples: VSCMount.exe --dl C --mp C:\VssRoot --debug

            Links are created in <mp>_<drive letter> (C:\VssRoot_C above). Use --ud false for names without dates
            Short options (single letter) are prefixed with a single dash. Long commands are prefixed with two dashes

    Usage:
      VSCMount [options]

    Options:
      --dl <dl>       Source drive to look for Volume Shadow Copies (C, D:, or F:\ for example)
      --mp <mp>       The base directory where you want VSCs mapped to
      --ud            Use VSC creation timestamps (yyyyMMddTHHmmss.fffffff, UTC) in symbolic link names. Default is true. Use --ud false to turn off
      --debug         Show debug information during processing
      --trace         Show trace information during processing
      -?, -h, --help  Show help and usage information
      --version       Show version information

## Documentation

Run VSCMount from an elevated (administrator) prompt on Windows. It needs .NET 10.

`VSCMount.exe --dl C --mp C:\VssRoot` creates `C:\VssRoot_C` with one symbolic link per shadow copy of C:, such as `vss001-20261017T170000.1234560` (the number of the shadow copy and when it was created, in UTC), or `vss001` with `--ud false`.

- Running it again replaces the links from the earlier run.
- The `<mp>_<drive letter>` folder must be new, empty or hold only VSCMount's links. If it holds anything else, VSCMount stops without deleting anything (versions before 2026.10.0 deleted everything in it).
- To remove access to the shadow copies, delete the links or the whole `<mp>_<drive letter>` folder. That removes the links only; the shadow copies are not changed.
- Exit codes: 0 on success (also when the drive has no shadow copies), 1 for invalid arguments, a drive that isn't ready, missing administrator rights, a folder that can't be used, or a shadow copy that failed to mount.

# Download Eric Zimmerman's Tools

All of Eric Zimmerman's tools can be downloaded [here](https://ericzimmerman.github.io/#!index.md).

# Special Thanks

Open Source Development funding and support provided by the following contributors:
- [SANS Institute](http://sans.org/) and [SANS DFIR](http://dfir.sans.org/).
- [Tines](https://www.tines.com/?utm_source=oss&utm_medium=sponsorship&utm_campaign=ericzimmerman)
