# VSCMount

Mounts every Volume Shadow Copy (VSC) of a drive as a folder: it creates a directory symbolic link for each shadow
copy, so you can browse it with Explorer or point other tools at it.

| | |
|---|---|
| **Version** | 2026.10.0 |
| **Download** | [.NET 10](https://download.ericzimmermanstools.com/net/VSCMount.zip) |
| **All tools** | [ericzimmerman.github.io](https://ericzimmerman.github.io) or [Get-ZimmermanTools](https://download.ericzimmermanstools.com/Get-ZimmermanTools.zip), which downloads them all and keeps them current |
| **Changes** | [ChangeLog](https://ericzimmerman.github.io/changelog/) |
| **Issues** | [EricZimmerman/VSCMount](https://github.com/EricZimmerman/VSCMount/issues) or [EricZimmerman/Issues](https://github.com/EricZimmerman/Issues/issues) |

## Requirements

- **.NET 10 build**: the [.NET 10 runtime](https://dotnet.microsoft.com/download/dotnet/10.0). Run `VSCMount.exe` on
  Windows, from an elevated (administrator) prompt: only administrators can query the shadow copies. `--help` works
  without administrator rights.

## Command line

```text
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
```

Short options (a single letter) start with one dash, long options with two.

## Examples

```text
VSCMount.exe --dl C --mp C:\VssRoot --debug
```

## Output

`VSCMount.exe --dl C --mp C:\VssRoot` creates `C:\VssRoot_C` (`<mp>_<drive letter>`) with one symbolic link per shadow
copy of C:, such as `vss001-20261017T170000.1234560` (the number of the shadow copy and when it was created, in UTC), or
`vss001` with `--ud false`. Each link points to `\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopyN\`.

- Running it again replaces the links from the earlier run.
- The `<mp>_<drive letter>` folder must be new, empty or hold only VSCMount's links. If it holds anything else,
  VSCMount stops without deleting anything (versions before 2026.10.0 deleted everything in it).
- A shadow copy whose volume WMI doesn't list is skipped with a warning; a volume mounted in a folder (with no drive
  letter) is never matched.
- To remove access to the shadow copies, delete the links or the whole `<mp>_<drive letter>` folder. That removes the
  links only; the shadow copies are not changed.

## Exit codes

| Code | Meaning |
|---|---|
| `0` | Success, also when the drive has no shadow copies |
| `1` | The tool failed: missing or invalid arguments, not Windows, no administrator rights, a drive that isn't ready, a mount folder that can't be used, a WMI error, or a shadow copy that failed to mount |

## What's new in 2026.10

- Only the `vss*` links of an earlier run are removed from the mount folder (`<mp>_<drive letter>`); a folder with
  anything else in it is left alone with an error. It was deleted with everything in it.
- `--help` works without administrator rights.
- A VSC on a volume that WMI doesn't list is skipped with a warning instead of stopping the run; volumes mounted in a
  folder no longer match a drive letter.
- VSC creation times use the UTC offset WMI gives instead of the local time zone.
- A VSC that fails to mount says why, and a failed mount can no longer be reported as mounted.
- No longer needs `WmiLight.Native.dll`, so it can't fail with "Unable to load DLL 'WmiLight.Native.dll'".
- Exit code 1 when the tool fails (it was always 0). Error reports include source file and line numbers.

The full list is in the [ChangeLog](https://ericzimmerman.github.io/changelog/).

## Building

```text
dotnet build VSCMount.slnx -c Release
```

Builds `net10.0` (Windows only at run time). WMI queries use the [WmiLight](https://www.nuget.org/packages/WmiLight)
library.

## Special thanks

Open source development funding and support provided by the [SANS Institute](https://www.sans.org/) and
[SANS DFIR](https://www.sans.org/digital-forensics-incident-response/).

## License

[MIT](LICENSE), Copyright (c) Eric R. Zimmerman 2027.
