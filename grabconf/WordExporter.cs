using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;

namespace GrabConf;

public sealed class WordExporter
{
    public void Export(
        string outputPath,
        string title,
        string spaceName,
        string htmlContent,
        List<DownloadedAttachment> attachments,
        PageMetadata metadata)
    {
        Log.Debug($"Processing HTML images for '{title}'...");
        var processedHtml = ProcessHtmlImages(htmlContent, attachments);

        Log.Debug($"Writing .docx to {outputPath}...");
        using var doc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        AddMetadataHeader(body, title, spaceName, metadata);

        body.AppendChild(CreateHorizontalRule());

        Log.Debug($"Converting HTML body ({processedHtml.Length:N0} chars) to Open XML...");
        var converter = new HtmlConverter(mainPart);
        var paragraphs = converter.Parse(processedHtml);
        foreach (var element in paragraphs)
            body.AppendChild(element);

        var nonImageAttachments = attachments
            .Where(a => !a.MediaType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (nonImageAttachments.Count > 0)
        {
            var folderName = Path.GetFileNameWithoutExtension(outputPath) + "_att";
            AddAttachmentsSection(body, nonImageAttachments, folderName);
        }

        mainPart.Document.Save();

        SaveAttachments(outputPath, attachments);
    }

    private static void AddMetadataHeader(Body body, string title, string spaceName, PageMetadata metadata)
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

        AddMetadataLine(body, "Space", spaceName);
        AddMetadataLine(body, "Title", title);
        AddMetadataLine(body, "Creator", creator);
        AddMetadataLine(body, "Contributors", contributors);
        AddMetadataLine(body, "Created", createdDate);
        AddMetadataLine(body, "Updates", updates);
        AddMetadataLine(body, "Last Updated", lastUpdated);
        AddMetadataLine(body, "Views", views);
        AddMetadataLine(body, "Exported", exportDate);
    }

    private static void AddMetadataLine(Body body, string label, string value)
    {
        var paragraph = new Paragraph();
        paragraph.AppendChild(new Run(
            new RunProperties(new Bold()),
            new Text(label + ": ") { Space = SpaceProcessingModeValues.Preserve }));
        paragraph.AppendChild(new Run(
            new Text(value)));
        body.AppendChild(paragraph);
    }

    private static Paragraph CreateHorizontalRule()
    {
        return new Paragraph(
            new ParagraphProperties(
                new ParagraphBorders(
                    new BottomBorder
                    {
                        Val = BorderValues.Single,
                        Size = 6,
                        Space = 1,
                        Color = "auto"
                    })));
    }

    private static void AddAttachmentsSection(Body body, List<DownloadedAttachment> attachments, string folderName)
    {
        body.AppendChild(CreateHorizontalRule());

        body.AppendChild(new Paragraph(
            new ParagraphProperties(
                new ParagraphStyleId { Val = "Heading2" }),
            new Run(new Text("Attachments"))));

        foreach (var att in attachments)
        {
            body.AppendChild(new Paragraph(
                new ParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference { Val = 0 },
                        new NumberingId { Val = 1 })),
                new Run(new Text(att.FileName))));
        }

        body.AppendChild(new Paragraph(
            new Run(
                new RunProperties(new Italic()),
                new Text($"Files saved to: {folderName}/"))));
    }

    private static string ProcessHtmlImages(string html, List<DownloadedAttachment> attachments)
    {
        var imageAttachments = attachments.Where(
            a => a.MediaType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)).ToList();
        Log.Debug($"Embedding {imageAttachments.Count} image(s) as base64 data URIs...");

        foreach (var att in imageAttachments)
        {
            var base64 = Convert.ToBase64String(att.Data);
            var dataUri = $"data:{att.MediaType};base64,{base64}";

            var escaped = Regex.Escape(att.FileName);
            var pattern = $@"src=""[^""]*?{escaped}[^""]*""";
            html = Regex.Replace(html, pattern, $@"src=""{dataUri}""", RegexOptions.IgnoreCase);
            Log.Debug($"Embedded image: {att.FileName} ({att.Data.Length:N0} bytes)");

            var encodedName = Uri.EscapeDataString(att.FileName);
            if (encodedName != att.FileName)
            {
                var encodedPattern = $@"src=""[^""]*?{Regex.Escape(encodedName)}[^""]*""";
                html = Regex.Replace(html, encodedPattern, $@"src=""{dataUri}""", RegexOptions.IgnoreCase);
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
}
