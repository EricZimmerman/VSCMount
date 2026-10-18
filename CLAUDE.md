# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

VSCMount is a Windows-only command line tool (`net10.0`; the net4 build was dropped in 2026.5) that creates a directory symbolic link for each Volume Shadow Copy of a drive. It is not packed; the user builds, signs and releases it. It needs an elevated prompt: `Win32_ShadowCopy` can only be queried by administrators.

- `Program.cs`: the CLI (`--dl`, `--mp`, `--ud`, `--debug`, `--trace`). Order of checks: required arguments, drive letter, Windows, administrator, drive ready, mount folder, WMI. `--help` and argument errors work without administrator rights (2026.5 checked for administrator before parsing).
- `VssParsing`: everything that doesn't need Windows, so it is tested on Linux CI. `ParseCimDateTime` (WMI `yyyyMMddHHmmss.ffffff+UUU`, the offset in minutes is used; 2026.5 dropped it and assumed local time), `GetShadowCopyNumber` (from `DeviceObject`), `IsShadowCopyTarget`, `GetLinkName`, `NormalizeDriveLetter` (`C`, `C:`, `C:\`, `C:/`; anything else is refused), `GetMountRoot`. None of them throw.
- `MountRoot.Prepare`: creates `<mp>_<letter>`, or removes the `vss*` links an earlier run left in it. If the folder holds anything else (a file, a real folder, a link to something that isn't a shadow copy), it deletes nothing and fails. 2026.5 deleted everything in that folder recursively, then the folder itself.
- `Helpers.GetVssInfoViaWmi`: WmiLight queries of `Win32_Volume` and `Win32_ShadowCopy`. A copy whose volume `Win32_Volume` doesn't list, or whose device name isn't `...HarddiskVolumeShadowCopyN`, is skipped with a warning (2026.5 threw and stopped). Volumes are matched by `Caption` (`C:\`); a volume mounted in a folder has no letter and is never matched (2026.5 matched it by the first character of its path).
- `Helpers.MountVss`: `CreateSymbolicLinkW` to `\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopyN\`. Its return type is `BOOLEAN`, so it is marshaled as `I1`; `SetLastError` gives the reason for a failure.

## Behavior to keep

- Links go in `<mp>_<drive letter>` (`C:\VssRoot` and `C:\VssRoot\` give `C:\VssRoot_C`) and are named `vss001` or, by default (`--ud` is true; `--ud false` turns it off), `vss001-yyyyMMddTHHmmss.fffffff` with the creation time in UTC.
- Never delete anything but VSCMount's own links: a `vss*` reparse point whose target is a shadow copy device. On Windows, `LinkTarget` reports such a link as `GLOBALROOT\Device\...` (without `\\?\`), so `IsShadowCopyTarget` accepts the target with or without the `\\?\`, `\??\` or `\\.\` prefix.
- Exit codes: 0 on success, also when the drive has no shadow copies. 1 for missing or invalid arguments, not Windows, no administrator rights, a drive that isn't ready, a mount folder that can't be used, a WMI error or any copy that failed to mount. No `Environment.Exit`.

## Commands

The SDK is pinned in `global.json` (.NET 10). `LangVersion 10`. Tests target `net10.0` and use NUnit 5 + FluentAssertions 7 (pinned to `[7.2.2,8.0.0)`, since 8.x requires a commercial license).

```bash
dotnet build -c Release                                           # VSCMount.slnx
dotnet test -c Release                                            # CI (test.yml); MountTests is [Explicit]
VSCMOUNT_FUZZ_ROUNDS=200000 dotnet test -c Release --filter Name~RandomInputNeverThrows
# From an elevated prompt on Windows, with at least one shadow copy of C: (VSCMOUNT_DRIVE for another drive):
dotnet test -c Release --filter "FullyQualifiedName~MountTests"
```

To make a shadow copy for MountTests (elevated PowerShell): `Invoke-CimMethod -ClassName Win32_ShadowCopy -MethodName Create -Arguments @{ Volume = 'C:\' }`.

## Tests

- `ParsingTests`: WMI dates, device names, link targets, link names, drive letters, mount folder names, and random input.
- `MountRootTests`: a real temp folder. Files, real `vss*` folders and `vss*` links to other places survive and make `Prepare` fail; links to shadow copy devices (dangling, made with `Directory.CreateSymbolicLink`) are removed. The link tests are ignored where links can't be created (Windows without administrator rights or Developer Mode).
- `CliTests`: VSCMount.dll as a process: help without administrator rights, exit codes for missing or invalid arguments, and that nothing is created without administrator rights (skipped when elevated, since it would mount).
- `MountTests` (`[Explicit]`, elevated Windows only): mounts the real shadow copies, checks names, targets and that the links show files; remounts; and checks a folder with other files is refused and kept.

## Versioning and release

- `<Version>` in the csproj, bumped by hand. The user builds, signs and releases it.
- `.github/workflows/test.yml` builds and tests on pushes to `master` and PRs (ubuntu; the Windows-only parts are covered by `MountTests`). Dependabot (`.github/dependabot.yml`) opens monthly PRs for GitHub Actions and NuGet updates.
