using System;
using FluentAssertions;
using NUnit.Framework;

namespace VSCMount.Test;

[TestFixture]
public class ParsingTests
{
    [TestCase("20261017120000.123456-300", "2026-10-17T17:00:00.1234560Z")]
    [TestCase("20260117120000.000000+060", "2026-01-17T11:00:00.0000000Z")]
    [TestCase("20260117120000.000000+000", "2026-01-17T12:00:00.0000000Z")]
    [TestCase("20260630233000.500000+330", "2026-06-30T18:00:00.5000000Z")]
    [TestCase("20261231230000.000000-120", "2027-01-01T01:00:00.0000000Z")]
    public void ParsesWmiDatesWithTheirOffset(string value, string expectedUtc)
    {
        var parsed = VssParsing.ParseCimDateTime(value);

        parsed.Should().NotBeNull();
        parsed.Value.Offset.Should().Be(TimeSpan.Zero);
        parsed.Value.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ").Should().Be(expectedUtc);
    }

    [TestCase(null)]
    [TestCase("")]
    [TestCase("2026")]
    [TestCase("20261017120000.123456")]
    [TestCase("20261317120000.123456-300")]
    [TestCase("20261017120000.123456*300")]
    [TestCase("20261017120000.123456-abc")]
    [TestCase("20261017120000.123456+999")]
    [TestCase("20261017120000.123456- 30")]
    [TestCase("********************.******+***")]
    [TestCase("20261017120000.123456-3000")]
    public void RejectsMalformedWmiDates(string value)
    {
        VssParsing.ParseCimDateTime(value).Should().BeNull();
    }

    [TestCase(@"\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy12", 12)]
    [TestCase(@"\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1\", 1)]
    [TestCase(@"\\?\globalroot\device\harddiskvolumeshadowcopy7", 7)]
    [TestCase(@"\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy0042", 42)]
    public void GetsTheShadowCopyNumber(string deviceObject, int expected)
    {
        VssParsing.GetShadowCopyNumber(deviceObject).Should().Be(expected);
    }

    [TestCase(null)]
    [TestCase("")]
    [TestCase(@"\\?\GLOBALROOT\Device\HarddiskVolume3")]
    [TestCase(@"\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy")]
    [TestCase(@"\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy-1")]
    [TestCase(@"\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy12a")]
    [TestCase(@"\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy99999999999")]
    public void RejectsOtherDevices(string deviceObject)
    {
        VssParsing.GetShadowCopyNumber(deviceObject).Should().BeNull();
    }

    [TestCase(@"\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1\", true)]
    [TestCase(@"\\?\globalroot\device\HarddiskVolumeShadowCopy12", true)]
    [TestCase(@"GLOBALROOT\Device\HarddiskVolumeShadowCopy1\", true)]
    [TestCase(@"\??\GLOBALROOT\Device\HarddiskVolumeShadowCopy1\", true)]
    [TestCase(@"\\.\GLOBALROOT\Device\HarddiskVolumeShadowCopy3", true)]
    [TestCase(@"\\?\C:\GLOBALROOT\Device\HarddiskVolumeShadowCopy1", false)]
    [TestCase(@"C:\Evidence\GLOBALROOT\Device\HarddiskVolumeShadowCopy1", false)]
    [TestCase(@"C:\Windows", false)]
    [TestCase(@"\\?\GLOBALROOT\Device\HarddiskVolume3\", false)]
    [TestCase("", false)]
    [TestCase(null, false)]
    public void RecognizesShadowCopyTargets(string target, bool expected)
    {
        VssParsing.IsShadowCopyTarget(target).Should().Be(expected);
    }

    [Test]
    public void NamesLinks()
    {
        var created = new DateTimeOffset(2026, 10, 17, 12, 0, 0, TimeSpan.FromHours(-5)).AddTicks(1234560);

        VssParsing.GetLinkName(1, created, true).Should().Be("vss001-20261017T170000.1234560");
        VssParsing.GetLinkName(12, created, false).Should().Be("vss012");
        VssParsing.GetLinkName(1, null, true).Should().Be("vss001");
        VssParsing.GetLinkName(1234, null, false).Should().Be("vss1234");
    }

    [TestCase("C", 'C')]
    [TestCase("c", 'C')]
    [TestCase("C:", 'C')]
    [TestCase(@"c:\", 'C')]
    [TestCase("C:/", 'C')]
    [TestCase(" D: ", 'D')]
    [TestCase("z", 'Z')]
    public void NormalizesDriveLetters(string value, char expected)
    {
        VssParsing.NormalizeDriveLetter(value).Should().Be(expected);
    }

    [TestCase(null)]
    [TestCase("")]
    [TestCase(" ")]
    [TestCase("1")]
    [TestCase("CC")]
    [TestCase(@"C:\x")]
    [TestCase("C;")]
    [TestCase(@"C:|")]
    [TestCase(@"\\?\")]
    [TestCase("é")]
    public void RejectsWhatIsNotADriveLetter(string value)
    {
        VssParsing.NormalizeDriveLetter(value).Should().BeNull();
    }

    [TestCase(@"C:\VssRoot", @"C:\VssRoot_C")]
    [TestCase(@"C:\VssRoot\", @"C:\VssRoot_C")]
    [TestCase(@"C:\VssRoot\\", @"C:\VssRoot_C")]
    [TestCase(@"C:\", @"C:\_C")]
    [TestCase(@"\\server\share\cases\", @"\\server\share\cases_C")]
    [TestCase("/mnt/vss/", "/mnt/vss_C")]
    [TestCase("/", "/_C")]
    public void BuildsTheMountRoot(string mountPoint, string expected)
    {
        VssParsing.GetMountRoot(mountPoint, 'C').Should().Be(expected);
    }

    [Test]
    public void RandomInputNeverThrows()
    {
        var rounds = int.TryParse(Environment.GetEnvironmentVariable("VSCMOUNT_FUZZ_ROUNDS"), out var r) ? r : 5000;
        var random = new Random(20261018);
        const string alphabet = @"0123456789.+-*:\/?CcGLOBALROOTHarddiskVolumeShadowCopy é";

        for (var i = 0; i < rounds; i++)
        {
            var chars = new char[random.Next(0, 60)];
            for (var j = 0; j < chars.Length; j++)
            {
                chars[j] = random.Next(10) == 0 ? (char)random.Next(0, 0xFFFF) : alphabet[random.Next(alphabet.Length)];
            }

            var value = new string(chars);

            //A valid prefix makes the parsers go further before they give up
            foreach (var input in new[] { value, "20261017120000.123456" + value, @"\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy" + value })
            {
                FluentActions.Invoking(() => VssParsing.ParseCimDateTime(input)).Should().NotThrow(input);
                FluentActions.Invoking(() => VssParsing.GetShadowCopyNumber(input)).Should().NotThrow(input);
                FluentActions.Invoking(() => VssParsing.IsShadowCopyTarget(input)).Should().NotThrow(input);
                FluentActions.Invoking(() => VssParsing.NormalizeDriveLetter(input)).Should().NotThrow(input);
                FluentActions.Invoking(() => VssParsing.GetMountRoot(input, 'C')).Should().NotThrow(input);
            }
        }
    }
}
