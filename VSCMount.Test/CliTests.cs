using System.Diagnostics;
using System.IO;
using FluentAssertions;
using NUnit.Framework;

namespace VSCMount.Test;

/// <summary>
///     Runs the built VSCMount.dll (copied next to the tests by the project reference) as a separate process. None of
///     these mount anything
/// </summary>
[TestFixture]
public class CliTests
{
    private string _dir;

    [SetUp]
    public void SetUp()
    {
        _dir = TestData.NewTempDir();
    }

    [TearDown]
    public void TearDown()
    {
        TestData.DeleteDir(_dir);
    }

    internal static (int ExitCode, string Output) Run(string workingDirectory, params string[] args)
    {
        var start = new ProcessStartInfo("dotnet")
        {
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            WorkingDirectory = workingDirectory
        };
        start.ArgumentList.Add(Path.Combine(TestContext.CurrentContext.TestDirectory, "VSCMount.dll"));
        foreach (var arg in args)
        {
            start.ArgumentList.Add(arg);
        }

        using var process = Process.Start(start)!;
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        process.StandardInput.Close();

        process.WaitForExit(60_000).Should().BeTrue("VSCMount should finish");

        return (process.ExitCode, stdout.Result + stderr.Result);
    }

    [Test]
    public void HelpWorksWithoutAdministratorRights()
    {
        var (exitCode, output) = Run(_dir, "--help");

        exitCode.Should().Be(0, output);
        output.Should().Contain("--dl").And.Contain("--mp").And.Contain("--ud");
        output.Should().NotContain("Administrator privileges not found");
    }

    [Test]
    public void MissingArgumentsFail()
    {
        var (exitCode, output) = Run(_dir);

        exitCode.Should().Be(1, output);
        output.Should().Contain("Both --dl and --mp are required");
    }

    [Test]
    public void MissingMountPointFails()
    {
        var (exitCode, output) = Run(_dir, "--dl", "C");

        exitCode.Should().Be(1, output);
        output.Should().Contain("Both --dl and --mp are required");
    }

    [TestCase("1")]
    [TestCase("CC")]
    [TestCase(@"C:\Windows")]
    public void InvalidDriveLetterFails(string dl)
    {
        var (exitCode, output) = Run(_dir, "--dl", dl, "--mp", Path.Combine(_dir, "VssRoot"));

        exitCode.Should().Be(1, output);
        output.Should().Contain("is not a drive letter");
    }

    [Test]
    public void UnknownOptionFails()
    {
        var (exitCode, output) = Run(_dir, "--dl", "C", "--mp", Path.Combine(_dir, "VssRoot"), "--nope");

        exitCode.Should().Be(1, output);
    }

    [Test]
    public void WithoutAdministratorRightsOrWindowsNothingIsCreated()
    {
        Assume.That(TestData.IsWindowsAdministrator, Is.False, "This would mount for real; see MountTests");

        var mp = Path.Combine(_dir, "VssRoot");
        var (exitCode, output) = Run(_dir, "--dl", "C", "--mp", mp);

        exitCode.Should().Be(1, output);
        output.Should().ContainAny("Administrator privileges not found", "only supported on Windows");
        Directory.Exists(mp + "_C").Should().BeFalse();
    }
}
