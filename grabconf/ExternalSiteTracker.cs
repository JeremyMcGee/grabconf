using System.Text.RegularExpressions;

namespace GrabConf;

public sealed partial class ExternalSiteTracker
{
    private readonly string _currentSiteBase;
    private readonly Dictionary<string, int> _siteCounts = new(StringComparer.OrdinalIgnoreCase);

    public bool HasReferences => _siteCounts.Count > 0;

    public ExternalSiteTracker(string currentBaseUrl)
    {
        _currentSiteBase = NormaliseBase(currentBaseUrl);
    }

    public void Scan(string html)
    {
        foreach (Match match in HrefRegex().Matches(html))
        {
            var url = match.Groups[1].Value;
            var siteBase = ExtractConfluenceSiteBase(url);
            if (siteBase is null)
                continue;

            if (siteBase.Equals(_currentSiteBase, StringComparison.OrdinalIgnoreCase))
                continue;

            if (_siteCounts.TryGetValue(siteBase, out var count))
                _siteCounts[siteBase] = count + 1;
            else
                _siteCounts[siteBase] = 1;
        }
    }

    public void WriteManifest(string outputPath, string sourceBaseUrl, string spaceKey)
    {
        var totalRefs = 0;
        foreach (var count in _siteCounts.Values)
            totalRefs += count;

        using var writer = new StreamWriter(outputPath, append: false, encoding: System.Text.Encoding.UTF8);
        writer.WriteLine("External Confluence Sites Manifest");
        writer.WriteLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        writer.WriteLine($"Source:    {sourceBaseUrl} (space: {spaceKey})");
        writer.WriteLine("---");
        writer.WriteLine();

        if (_siteCounts.Count == 0)
        {
            writer.WriteLine("No external Confluence site references found.");
        }
        else
        {
            var sorted = _siteCounts.OrderByDescending(kv => kv.Value);
            var maxUrlLen = _siteCounts.Keys.Max(k => k.Length);

            foreach (var (site, count) in sorted)
                writer.WriteLine($"{site.PadRight(maxUrlLen)}  {count}");

            writer.WriteLine();
            writer.WriteLine($"Total: {totalRefs} reference(s) to {_siteCounts.Count} external site(s)");
        }
    }

    private static string? ExtractConfluenceSiteBase(string url)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
            return null;

        if (uri.Scheme is not ("http" or "https"))
            return null;

        var path = uri.AbsolutePath;
        var origin = $"{uri.Scheme}://{uri.Authority}";

        if (path.StartsWith("/wiki/", StringComparison.OrdinalIgnoreCase) ||
            path.Equals("/wiki", StringComparison.OrdinalIgnoreCase))
            return $"{origin}/wiki";

        if (path.StartsWith("/confluence/", StringComparison.OrdinalIgnoreCase) ||
            path.Equals("/confluence", StringComparison.OrdinalIgnoreCase))
            return $"{origin}/confluence";

        if (path.StartsWith("/display/", StringComparison.OrdinalIgnoreCase) ||
            path.StartsWith("/pages/viewpage.action", StringComparison.OrdinalIgnoreCase) ||
            path.StartsWith("/x/", StringComparison.OrdinalIgnoreCase))
            return origin;

        if (uri.Host.EndsWith(".atlassian.net", StringComparison.OrdinalIgnoreCase))
            return $"{origin}/wiki";

        return null;
    }

    private static string NormaliseBase(string baseUrl)
    {
        if (Uri.TryCreate(baseUrl, UriKind.Absolute, out var uri))
            return $"{uri.Scheme}://{uri.Authority}{uri.AbsolutePath.TrimEnd('/')}";

        return baseUrl.TrimEnd('/');
    }

    [GeneratedRegex(@"href=""(https?://[^""]+)""", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
    private static partial Regex HrefRegex();
}
