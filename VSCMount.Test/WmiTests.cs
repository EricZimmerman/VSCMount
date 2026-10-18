using System;
using FluentAssertions;
using NUnit.Framework;

namespace VSCMount.Test;

/// <summary>
///     WmiLight with its native dll found by Helpers' resolver. Win32_Volume needs no administrator rights
/// </summary>
[TestFixture]
public class WmiTests
{
    [Test]
    public void ListsVolumes()
    {
        Assume.That(OperatingSystem.IsWindows(), "WMI is Windows only");

        var volumes = Helpers.GetVolumes();

        volumes.Should().NotBeEmpty();
        volumes.Values.Should().Contain(t => VssParsing.NormalizeDriveLetter(t) != null, "the system drive has a letter");
    }
}
