using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;

namespace GrabConf;

public sealed class HtmlExporter
{
    private static readonly HashSet<string> SemanticTags = new(StringComparer.OrdinalIgnoreCase)
    {
        "h1", "h2", "h3", "h4", "h5", "h6",
        "p", "br", "hr",
        "ul", "ol", "li",
        "table", "thead", "tbody", "tfoot", "tr", "td", "th", "caption", "colgroup", "col",
        "a", "img",
        "blockquote", "pre", "code",
        "strong", "b", "em", "i", "u", "s", "del", "sup", "sub",
        "dl", "dt", "dd",
        "figure", "figcaption",
        "details", "summary",
        "abbr", "cite", "mark", "time",
    };

    private static readonly Dictionary<string, HashSet<string>> SafeAttributes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["a"] = new(StringComparer.OrdinalIgnoreCase) { "href", "title" },
        ["img"] = new(StringComparer.OrdinalIgnoreCase) { "src", "alt", "title", "width", "height" },
        ["td"] = new(StringComparer.OrdinalIgnoreCase) { "colspan", "rowspan" },
        ["th"] = new(StringComparer.OrdinalIgnoreCase) { "colspan", "rowspan", "scope" },
        ["col"] = new(StringComparer.OrdinalIgnoreCase) { "span" },
        ["colgroup"] = new(StringComparer.OrdinalIgnoreCase) { "span" },
        ["time"] = new(StringComparer.OrdinalIgnoreCase) { "datetime" },
        ["abbr"] = new(StringComparer.OrdinalIgnoreCase) { "title" },
        ["ol"] = new(StringComparer.OrdinalIgnoreCase) { "start", "type" },
    };

    public void Export(
        string outputPath,
        string title,
        string spaceName,
        string htmlContent,
        List<DownloadedAttachment> attachments,
        PageMetadata metadata)
    {
        Log.Debug($"Building HTML document for '{title}'...");

        var processedHtml = RewriteImageSources(htmlContent, attachments, outputPath);

        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang=\"en\">");

        AppendHead(sb, title, spaceName, metadata);

        sb.AppendLine("<body>");

        AppendMetadataHeader(sb, title, spaceName, metadata);

        sb.AppendLine("<main>");
        sb.AppendLine(CleanHtml(processedHtml));
        sb.AppendLine("</main>");

        var nonImageAttachments = attachments
            .Where(a => !a.MediaType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (nonImageAttachments.Count > 0)
        {
            var folderName = Path.GetFileNameWithoutExtension(outputPath) + "_att";
            AppendAttachmentsSection(sb, nonImageAttachments, folderName);
        }

        sb.AppendLine("</body>");
        sb.AppendLine("</html>");

        Log.Debug($"Writing HTML to {outputPath}...");
        File.WriteAllText(outputPath, sb.ToString(), Encoding.UTF8);

        SaveAttachments(outputPath, attachments);
    }

    public static void GenerateIndexFiles(string rootDir, string spaceName, Dictionary<string, string> pageTitles)
    {
        Log.Debug("Generating index files...");
        WriteIndex(rootDir, rootDir, spaceName, pageTitles);
    }

    private static void WriteIndex(string dir, string rootDir, string spaceName, Dictionary<string, string> pageTitles)
    {
        var subDirs = Directory.GetDirectories(dir)
            .Where(d => !Path.GetFileName(d).EndsWith("_att", StringComparison.OrdinalIgnoreCase))
            .OrderBy(d => d)
            .ToList();

        foreach (var subDir in subDirs)
            WriteIndex(subDir, rootDir, spaceName, pageTitles);

        var isRoot = Path.GetFullPath(dir).Equals(Path.GetFullPath(rootDir), StringComparison.OrdinalIgnoreCase);
        var heading = isRoot ? spaceName : Path.GetFileName(dir);

        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang=\"en\">");
        sb.AppendLine("<head>");
        sb.AppendLine("<meta charset=\"utf-8\">");
        sb.AppendLine($"<title>{Encode(heading)}</title>");
        sb.AppendLine("<meta name=\"generator\" content=\"grabconf\">");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine($"<h1>{Encode(heading)}</h1>");
        sb.AppendLine("<nav>");
        sb.AppendLine("<ul>");

        foreach (var subDir in subDirs)
        {
            var subName = Path.GetFileName(subDir);
            sb.AppendLine($"<li>&#128193; <a href=\"{Uri.EscapeDataString(subName)}/index.html\">{Encode(subName)}</a></li>");
        }

        var htmlFiles = Directory.GetFiles(dir, "*.html")
            .Where(f => !Path.GetFileName(f).Equals("index.html", StringComparison.OrdinalIgnoreCase))
            .OrderBy(f => f)
            .ToList();

        foreach (var file in htmlFiles)
        {
            var fileName = Path.GetFileName(file);
            var fullPath = Path.GetFullPath(file);
            var displayName = pageTitles.TryGetValue(fullPath, out var title)
                ? title
                : Path.GetFileNameWithoutExtension(fileName);
            sb.AppendLine($"<li>&#128196; <a href=\"{Uri.EscapeDataString(fileName)}\">{Encode(displayName)}</a></li>");
        }

        sb.AppendLine("</ul>");
        sb.AppendLine("</nav>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");

        var indexPath = Path.Combine(dir, "index.html");
        File.WriteAllText(indexPath, sb.ToString(), Encoding.UTF8);
        Log.Debug($"Generated index: {indexPath}");
    }

    private static void AppendHead(StringBuilder sb, string title, string spaceName, PageMetadata metadata)
    {
        var creator = metadata.CreatorName ?? "Unknown";
        var contributors = metadata.Contributors.Count > 0
            ? string.Join(", ", metadata.Contributors)
            : creator;
        var description = $"Confluence page from space '{spaceName}' — version {metadata.VersionNumber}";

        sb.AppendLine("<head>");
        sb.AppendLine("<meta charset=\"utf-8\">");
        sb.AppendLine($"<title>{Encode(title)} — {Encode(spaceName)}</title>");
        sb.AppendLine($"<meta name=\"author\" content=\"{Encode(creator)}\">");
        sb.AppendLine($"<meta name=\"description\" content=\"{Encode(description)}\">");
        sb.AppendLine($"<meta name=\"keywords\" content=\"Confluence, {Encode(spaceName)}, {Encode(title)}\">");
        if (metadata.CreatedDate.HasValue)
            sb.AppendLine($"<meta name=\"date\" content=\"{metadata.CreatedDate.Value.UtcDateTime:yyyy-MM-dd}\">");
        if (metadata.LastUpdatedDate.HasValue)
            sb.AppendLine($"<meta name=\"last-modified\" content=\"{metadata.LastUpdatedDate.Value.UtcDateTime:yyyy-MM-dd}\">");
        sb.AppendLine($"<meta name=\"generator\" content=\"grabconf\">");
        sb.AppendLine("</head>");
    }

    private static void AppendMetadataHeader(StringBuilder sb, string title, string spaceName, PageMetadata metadata)
    {
        var exportDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        var creator = metadata.CreatorName ?? "Unknown";
        var contributors = metadata.Contributors.Count > 0
            ? string.Join(", ", metadata.Contributors)
            : "None";
        var createdDate = metadata.CreatedDate?.LocalDateTime.ToString("yyyy-MM-dd HH:mm:ss") ?? "Unknown";
        var updates = metadata.VersionNumber.ToString();
        var lastUpdated = metadata.LastUpdatedDate?.LocalDateTime.ToString("yyyy-MM-dd HH:mm:ss") ?? "Unknown";
        var views = metadata.ViewCount?.ToString("N0") ?? "N/A";

        sb.AppendLine("<header>");
        sb.AppendLine($"<h1>{Encode(title)}</h1>");
        sb.AppendLine("<dl>");
        AppendDt(sb, "Space", spaceName);
        AppendDt(sb, "Creator", creator);
        AppendDt(sb, "Contributors", contributors);
        AppendDt(sb, "Created", createdDate);
        AppendDt(sb, "Updates", updates);
        AppendDt(sb, "Last Updated", lastUpdated);
        AppendDt(sb, "Views", views);
        AppendDt(sb, "Exported", exportDate);
        sb.AppendLine("</dl>");
        sb.AppendLine("</header>");
    }

    private static void AppendDt(StringBuilder sb, string label, string value)
    {
        sb.AppendLine($"<dt>{Encode(label)}:</dt><dd>{Encode(value)}</dd>");
    }

    private static void AppendAttachmentsSection(StringBuilder sb, List<DownloadedAttachment> attachments, string folderName)
    {
        sb.AppendLine("<footer>");
        sb.AppendLine("<h2>Attachments</h2>");
        sb.AppendLine("<ul>");
        foreach (var att in attachments)
            sb.AppendLine($"<li><a href=\"{folderName}/{Uri.EscapeDataString(att.FileName)}\">{Encode(att.FileName)}</a></li>");
        sb.AppendLine("</ul>");
        sb.AppendLine($"<p><em>Files saved to: {Encode(folderName)}/</em></p>");
        sb.AppendLine("</footer>");
    }

    private static string RewriteImageSources(string html, List<DownloadedAttachment> attachments, string outputPath)
    {
        var imageAttachments = attachments
            .Where(a => a.MediaType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (imageAttachments.Count == 0)
            return html;

        var folderName = Path.GetFileNameWithoutExtension(outputPath) + "_att";
        Log.Debug($"Rewriting {imageAttachments.Count} image src(s) to local paths in {folderName}/...");

        foreach (var att in imageAttachments)
        {
            var relativePath = $"{folderName}/{Uri.EscapeDataString(att.FileName)}";

            var escaped = Regex.Escape(att.FileName);
            var pattern = $@"src=""[^""]*?{escaped}[^""]*""";
            html = Regex.Replace(html, pattern, $@"src=""{relativePath}""", RegexOptions.IgnoreCase);

            var encodedName = Uri.EscapeDataString(att.FileName);
            if (encodedName != att.FileName)
            {
                var encodedPattern = $@"src=""[^""]*?{Regex.Escape(encodedName)}[^""]*""";
                html = Regex.Replace(html, encodedPattern, $@"src=""{relativePath}""", RegexOptions.IgnoreCase);
            }
        }

        return html;
    }

    private static void SaveAttachments(string docPath, List<DownloadedAttachment> attachments)
    {
        if (attachments.Count == 0)
            return;

        var dir = Path.Combine(
            Path.GetDirectoryName(docPath) ?? ".",
            Path.GetFileNameWithoutExtension(docPath) + "_att");

        Directory.CreateDirectory(dir);
        Log.Debug($"Saving {attachments.Count} attachment(s) to {dir}");

        foreach (var att in attachments)
        {
            var filePath = Path.Combine(dir, att.FileName);
            File.WriteAllBytes(filePath, att.Data);
            Log.Debug($"Saved attachment: {filePath} ({att.Data.Length:N0} bytes)");
        }
    }

    private static string Encode(string value) => WebUtility.HtmlEncode(value);

    private static string CleanHtml(string rawHtml)
    {
        if (string.IsNullOrWhiteSpace(rawHtml))
            return string.Empty;

        var parser = new HtmlParser();
        var doc = parser.ParseDocument($"<body>{rawHtml}</body>");
        var body = doc.Body!;

        CleanElement(body);

        return body.InnerHtml.Trim();
    }

    private static void CleanElement(IElement element)
    {
        // Process children bottom-up so removals don't shift indices.
        var children = element.Children.ToList();
        foreach (var child in children)
            CleanElement(child);

        // Don't unwrap the synthetic <body> root.
        if (element.LocalName.Equals("body", StringComparison.OrdinalIgnoreCase))
            return;

        if (SemanticTags.Contains(element.LocalName))
        {
            // Keep the element but strip non-safe attributes.
            SafeAttributes.TryGetValue(element.LocalName, out var allowed);
            var attrs = element.Attributes.Select(a => a.Name).ToList();
            foreach (var attr in attrs)
            {
                if (allowed is null || !allowed.Contains(attr))
                    element.RemoveAttribute(attr);
            }
        }
        else
        {
            // Non-semantic element (div, span, etc.) — unwrap: replace with children.
            element.OuterHtml = element.InnerHtml;
        }
    }
}
