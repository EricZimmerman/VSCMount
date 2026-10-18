using System;
using System.Globalization;

namespace VSCMount;

/// <summary>
///     Parsing and naming that doesn't need Windows, so it can be tested anywhere. Nothing here throws on bad input
/// </summary>
public static class VssParsing
{
    private const string ShadowCopyDevice = "HarddiskVolumeShadowCopy";

    /// <summary>
    ///     Parses a WMI CIM_DATETIME value (<c>yyyyMMddHHmmss.ffffff+UUU</c>, where UUU is the offset from UTC in minutes)
    ///     and returns it in UTC. Null when the value is missing or malformed
    /// </summary>
    public static DateTimeOffset? ParseCimDateTime(string value)
    {
        if (value == null || value.Length != 25) return null;

        if (!DateTime.TryParseExact(value.Substring(0, 21), "yyyyMMddHHmmss.ffffff", CultureInfo.InvariantCulture,
                DateTimeStyles.None, out var dateTime))
            return null;

        var sign = value[21];
        if (sign != '+' && sign != '-') return null;

        if (!int.TryParse(value.Substring(22, 3), NumberStyles.None, CultureInfo.InvariantCulture, out var minutes))
            return null;

        //DateTimeOffset allows offsets up to 14 hours
        if (minutes > 14 * 60) return null;

        var offset = TimeSpan.FromMinutes(sign == '-' ? -minutes : minutes);

        return new DateTimeOffset(dateTime, offset).ToUniversalTime();
    }

    /// <summary>
    ///     The shadow copy number from a device object such as <c>\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy12</c>.
    ///     Null when it isn't a shadow copy device
    /// </summary>
    public static int? GetShadowCopyNumber(string deviceObject)
    {
        if (string.IsNullOrEmpty(deviceObject)) return null;

        var index = deviceObject.LastIndexOf(ShadowCopyDevice, StringComparison.OrdinalIgnoreCase);
        if (index < 0) return null;

        var digits = deviceObject.Substring(index + ShadowCopyDevice.Length).TrimEnd('\\');

        if (!int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out var number)) return null;

        return number;
    }

    /// <summary>
    ///     True when a link target points to a shadow copy device (what VSCMount creates). Links are created to
    ///     <c>\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopyN\</c>, but Windows reports the target with that prefix
    ///     removed (<c>GLOBALROOT\...</c>) or in NT form (<c>\??\GLOBALROOT\...</c>), so each of those matches
    /// </summary>
    public static bool IsShadowCopyTarget(string linkTarget)
    {
        if (linkTarget == null) return false;

        foreach (var prefix in new[] { @"\\?\", @"\??\", @"\\.\" })
        {
            if (linkTarget.StartsWith(prefix, StringComparison.Ordinal))
            {
                linkTarget = linkTarget.Substring(prefix.Length);
                break;
            }
        }

        return linkTarget.StartsWith(@"GLOBALROOT\Device\" + ShadowCopyDevice, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    ///     The link name for a shadow copy: <c>vss012</c>, or <c>vss012-20261017T120000.1234560</c> (UTC) with dates
    /// </summary>
    public static string GetLinkName(int vssNumber, DateTimeOffset? createdOn, bool useDatesInNames)
    {
        var name = $"vss{vssNumber.ToString("000", CultureInfo.InvariantCulture)}";

        if (useDatesInNames && createdOn.HasValue)
            name += $"-{createdOn.Value.UtcDateTime.ToString("yyyyMMddTHHmmss.fffffff", CultureInfo.InvariantCulture)}";

        return name;
    }

    /// <summary>
    ///     Normalizes <c>C</c>, <c>c:</c>, <c>C:\</c> or <c>C:/</c> to an upper case drive letter. Null for anything else
    /// </summary>
    public static char? NormalizeDriveLetter(string value)
    {
        if (value == null) return null;

        value = value.Trim();

        if (value.Length == 0 || value.Length > 3) return null;

        var letter = char.ToUpperInvariant(value[0]);
        if (letter < 'A' || letter > 'Z') return null;

        if (value.Length >= 2 && value[1] != ':') return null;
        if (value.Length == 3 && value[2] != '\\' && value[2] != '/') return null;

        return letter;
    }

    /// <summary>
    ///     The folder the links go in: <c>--mp</c> followed by <c>_</c> and the drive letter (<c>C:\VssRoot</c> and
    ///     <c>C:\VssRoot\</c> both give <c>C:\VssRoot_C</c>; <c>C:\</c> gives <c>C:\_C</c>)
    /// </summary>
    public static string GetMountRoot(string mountPoint, char driveLetter)
    {
        var trimmed = mountPoint.TrimEnd('\\', '/');

        //Keep the separator of a drive root (C:\) or a Unix root (/)
        if ((trimmed.Length == 0 || (trimmed.Length == 2 && trimmed[1] == ':')) && mountPoint.Length > trimmed.Length)
            trimmed = mountPoint.Substring(0, trimmed.Length + 1);

        return $"{trimmed}_{driveLetter}";
    }
}
