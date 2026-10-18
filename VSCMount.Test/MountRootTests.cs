using System.IO;
using FluentAssertions;
using NUnit.Framework;

namespace VSCMount.Test;

/// <summary>
///     2026.5 deleted everything in the mount folder recursively. Only VSCMount's own links may ever be removed
/// </summary>
[TestFixture]
public class MountRootTests
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

    private string Root => Path.Combine(_dir, "VssRoot_C");

    [Test]
    public void CreatesAMissingFolder()
    {
        MountRoot.Prepare(Root, out var error).Should().BeTrue(error);

        Directory.Exists(Root).Should().BeTrue();
    }

    [Test]
    public void UsesAnEmptyFolder()
    {
        Directory.CreateDirectory(Root);

        MountRoot.Prepare(Root, out var error).Should().BeTrue(error);
    }

    [Test]
    public void RefusesAFolderWithFilesAndKeepsThem()
    {
        Directory.CreateDirectory(Path.Combine(Root, "Case123"));
        File.WriteAllText(Path.Combine(Root, "Case123", "evidence.txt"), "keep me");
        File.WriteAllText(Path.Combine(Root, "notes.txt"), "keep me too");

        MountRoot.Prepare(Root, out var error).Should().BeFalse();

        error.Should().Contain("2 item(s)").And.Contain("Case123").And.Contain("notes.txt").And.Contain("Nothing was deleted");
        File.ReadAllText(Path.Combine(Root, "Case123", "evidence.txt")).Should().Be("keep me");
        File.ReadAllText(Path.Combine(Root, "notes.txt")).Should().Be("keep me too");
    }

    [Test]
    public void RefusesARealFolderNamedVss()
    {
        Directory.CreateDirectory(Path.Combine(Root, "vss001"));
        File.WriteAllText(Path.Combine(Root, "vss001", "evidence.txt"), "keep me");

        MountRoot.Prepare(Root, out var error).Should().BeFalse();

        error.Should().Contain("vss001");
        File.Exists(Path.Combine(Root, "vss001", "evidence.txt")).Should().BeTrue();
    }

    [Test]
    public void RefusesAFile()
    {
        File.WriteAllText(Root, "not a folder");

        MountRoot.Prepare(Root, out var error).Should().BeFalse();

        error.Should().Contain("is a file");
        File.Exists(Root).Should().BeTrue();
    }

    [Test]
    public void RemovesLinksFromAnEarlierRun()
    {
        Directory.CreateDirectory(Root);
        TestData.CreateLinkOrIgnore(Path.Combine(Root, "vss001"), TestData.ShadowCopyTarget);
        TestData.CreateLinkOrIgnore(Path.Combine(Root, "vss002-20261017T170000.1234560"),
            @"\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy2\");

        MountRoot.Prepare(Root, out var error).Should().BeTrue(error);

        Directory.GetFileSystemEntries(Root).Should().BeEmpty();
    }

    [Test]
    public void DeletesNothingWhenLinksAndOtherItemsAreMixed()
    {
        Directory.CreateDirectory(Root);
        TestData.CreateLinkOrIgnore(Path.Combine(Root, "vss001"), TestData.ShadowCopyTarget);
        File.WriteAllText(Path.Combine(Root, "notes.txt"), "keep me");

        MountRoot.Prepare(Root, out var error).Should().BeFalse();

        error.Should().Contain("1 item(s)").And.Contain("notes.txt");
        Directory.GetFileSystemEntries(Root).Should().HaveCount(2, "the link must not be removed either");
    }

    [Test]
    public void RefusesAVssLinkToSomethingElseAndKeepsItsTarget()
    {
        var target = Path.Combine(_dir, "Evidence");
        Directory.CreateDirectory(target);
        File.WriteAllText(Path.Combine(target, "evidence.txt"), "keep me");

        Directory.CreateDirectory(Root);
        TestData.CreateLinkOrIgnore(Path.Combine(Root, "vss001"), target);

        MountRoot.Prepare(Root, out var error).Should().BeFalse();

        error.Should().Contain("vss001");
        File.Exists(Path.Combine(target, "evidence.txt")).Should().BeTrue();
    }
}
