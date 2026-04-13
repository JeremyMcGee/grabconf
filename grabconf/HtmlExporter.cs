using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace GrabConf;

public sealed class HtmlExporter
{
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
        sb.AppendLine(processedHtml);
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
        sb.AppendLine("<style>");
        sb.AppendLine("  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; max-width: 960px; margin: 2em auto; padding: 0 1em; }");
        sb.AppendLine("  header { border-bottom: 1px solid #ccc; padding-bottom: 1em; margin-bottom: 2em; }");
        sb.AppendLine("  header h1 { margin-bottom: 0.25em; }");
        sb.AppendLine("  .metadata { font-size: 0.9em; color: #555; }");
        sb.AppendLine("  .metadata dt { font-weight: bold; display: inline; }");
        sb.AppendLine("  .metadata dd { display: inline; margin: 0 1.5em 0 0; }");
        sb.AppendLine("  footer { border-top: 1px solid #ccc; padding-top: 1em; margin-top: 2em; font-size: 0.9em; color: #555; }");
        sb.AppendLine("</style>");
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
        sb.AppendLine("<dl class=\"metadata\">");
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
}
