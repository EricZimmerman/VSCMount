using System;
using System.IO;
using System.Linq;
using FluentAssertions;
using NUnit.Framework;

namespace VSCMount.Test;

/// <summary>
///     Mounts the real shadow copies of C: (VSCMOUNT_DRIVE to use another drive). Needs Windows, an elevated prompt and
///     at least one shadow copy. To make one:
///     <c>Invoke-CimMethod -ClassName Win32_ShadowCopy -MethodName Create -Arguments @{ Volume = 'C:\' }</c>
/// </summary>
[TestFixture]
[Explicit("Needs Windows, administrator rights and shadow copies")]
public class MountTests
{
    private string _dir;
    private string _drive;

    [SetUp]
    public void SetUp()
    {
        if (!TestData.IsWindowsAdministrator) Assert.Ignore("Run from an elevated prompt on Windows");

        _drive = Environment.GetEnvironmentVariable("VSCMOUNT_DRIVE") ?? "C";
        _dir = TestData.NewTempDir();
    }

    [TearDown]
    public void TearDown()
    {
        if (_dir != null) TestData.DeleteDir(_dir);
    }

    private string MountPoint => Path.Combine(_dir, "VssRoot");
    private string Root => MountPoint + "_" + _drive;

    private string[] Links()
    {
        return Directory.GetDirectories(Root, "vss*").OrderBy(t => t).ToArray();
    }

    private void MountOrIgnore(params string[] extra)
    {
        var (exitCode, output) = CliTests.Run(_dir, new[] { "--dl", _drive, "--mp", MountPoint }.Concat(extra).ToArray());

        if (output.Contains("No VSCs found")) Assert.Ignore($"{_drive}: has no shadow copies");

        TestContext.Out.WriteLine(output);
        exitCode.Should().Be(0, output);
    }

    [Test]
    public void MountsEveryShadowCopy()
    {
        MountOrIgnore();

        var links = Links();
        links.Should().NotBeEmpty();

        foreach (var link in links)
        {
            var info = new DirectoryInfo(link);
            info.Attributes.Should().HaveFlag(FileAttributes.ReparsePoint);
            VssParsing.IsShadowCopyTarget(info.LinkTarget).Should().BeTrue(info.LinkTarget);
            Path.GetFileName(link).Should().MatchRegex(@"^vss\d{3,}-\d{8}T\d{6}\.\d{7}$");
            Directory.EnumerateFileSystemEntries(link).Should().NotBeEmpty($"{link} should show the volume's files");
        }
    }

    [Test]
    public void NamesWithoutDates()
    {
        MountOrIgnore("--ud", "false");

        Links().Select(Path.GetFileName).Should().OnlyContain(t => System.Text.RegularExpressions.Regex.IsMatch(t, @"^vss\d{3,}$"));
    }

    [Test]
    public void RemountsOverAnEarlierRun()
    {
        MountOrIgnore();
        var first = Links();

        MountOrIgnore();

        Links().Should().Equal(first);
    }

    [Test]
    public void LeavesAFolderWithOtherFilesAlone()
    {
        Directory.CreateDirectory(Root);
        File.WriteAllText(Path.Combine(Root, "evidence.txt"), "keep me");

        var (exitCode, output) = CliTests.Run(_dir, "--dl", _drive, "--mp", MountPoint);

        exitCode.Should().Be(1, output);
        output.Should().Contain("Nothing was deleted");
        File.ReadAllText(Path.Combine(Root, "evidence.txt")).Should().Be("keep me");
        Directory.GetDirectories(Root, "vss*").Should().BeEmpty();
    }
}
