using System.Text;
using System.Text.RegularExpressions;
using ReverseMarkdown;

namespace GrabConf;

public sealed class MarkdownExporter
{
    private static readonly Converter MarkdownConverter = new(new Config
    {
        GithubFlavored = true,
        RemoveComments = true,
        SmartHrefHandling = true,
        UnknownTags = Config.UnknownTagsOption.Bypass
    });

    public void Export(
        string outputPath,
        string title,
        string spaceName,
        string htmlContent,
        List<DownloadedAttachment> attachments,
        PageMetadata metadata)
    {
        Log.Debug($"Converting HTML to Markdown for '{title}'...");

        var sb = new StringBuilder();

        AppendMetadataHeader(sb, title, spaceName, metadata);

        sb.AppendLine("---");
        sb.AppendLine();

        var processedHtml = RewriteImageSources(htmlContent, attachments, outputPath);
        var markdown = MarkdownConverter.Convert(processedHtml);
        sb.AppendLine(markdown.TrimEnd());

        var nonImageAttachments = attachments
            .Where(a => !a.MediaType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (nonImageAttachments.Count > 0)
        {
            var folderName = Path.GetFileNameWithoutExtension(outputPath) + "_att";
            AppendAttachmentsSection(sb, nonImageAttachments, folderName);
        }

        Log.Debug($"Writing Markdown to {outputPath}...");
        File.WriteAllText(outputPath, sb.ToString(), Encoding.UTF8);

        SaveAttachments(outputPath, attachments);
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

        sb.AppendLine($"# {title}");
        sb.AppendLine();
        sb.AppendLine($"| Property | Value |");
        sb.AppendLine($"| --- | --- |");
        sb.AppendLine($"| **Space** | {EscapePipe(spaceName)} |");
        sb.AppendLine($"| **Creator** | {EscapePipe(creator)} |");
        sb.AppendLine($"| **Contributors** | {EscapePipe(contributors)} |");
        sb.AppendLine($"| **Created** | {createdDate} |");
        sb.AppendLine($"| **Updates** | {updates} |");
        sb.AppendLine($"| **Last Updated** | {lastUpdated} |");
        sb.AppendLine($"| **Views** | {views} |");
        sb.AppendLine($"| **Exported** | {exportDate} |");
        sb.AppendLine();
    }

    private static void AppendAttachmentsSection(StringBuilder sb, List<DownloadedAttachment> attachments, string folderName)
    {
        sb.AppendLine();
        sb.AppendLine("---");
        sb.AppendLine();
        sb.AppendLine("## Attachments");
        sb.AppendLine();

        foreach (var att in attachments)
            sb.AppendLine($"- [{att.FileName}]({folderName}/{Uri.EscapeDataString(att.FileName)})");

        sb.AppendLine();
        sb.AppendLine($"_Files saved to: {folderName}/_");
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

    private static string EscapePipe(string value) => value.Replace("|", "\\|");
}
