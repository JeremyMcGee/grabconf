using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GrabConf;

public sealed class WordExporter
{
    public void Export(
        string outputPath,
        string title,
        string spaceName,
        string htmlContent,
        List<DownloadedAttachment> attachments)
    {
        var processedHtml = ProcessHtmlImages(htmlContent, attachments);

        var nonImageAttachments = attachments
            .Where(a => !a.MediaType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
            .ToList();

        var attachmentsFolderName = Path.GetFileNameWithoutExtension(outputPath) + "_attachments";
        var attachmentsSection = BuildAttachmentsHtml(nonImageAttachments, attachmentsFolderName);

        var fullHtml = BuildFullHtml(title, spaceName, processedHtml, attachmentsSection);

        using var doc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();

        const string altChunkId = "pageContent";
        var chunk = mainPart.AddAlternativeFormatImportPart(
            AlternativeFormatImportPartType.Html, altChunkId);

        using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(fullHtml)))
        {
            chunk.FeedData(stream);
        }

        mainPart.Document = new Document(
            new Body(
                new AltChunk { Id = altChunkId }));

        mainPart.Document.Save();

        SaveAttachments(outputPath, attachments);
    }

    private static string BuildFullHtml(string title, string spaceName, string content, string attachmentsSection)
    {
        var encodedTitle = WebUtility.HtmlEncode(title);
        var encodedSpace = WebUtility.HtmlEncode(spaceName);
        var exportDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        return $"""
            <!DOCTYPE html>
            <html>
            <head>
              <meta charset="utf-8">
              <title>{encodedTitle}</title>
            </head>
            <body>
              <p><strong>Space:</strong> {encodedSpace}</p>
              <p><strong>Title:</strong> {encodedTitle}</p>
              <p><strong>Exported:</strong> {exportDate}</p>
              <hr>
              <h1>{encodedTitle}</h1>
              {content}
              {attachmentsSection}
            </body>
            </html>
            """;
    }

    private static string BuildAttachmentsHtml(
        List<DownloadedAttachment> nonImageAttachments,
        string folderName)
    {
        if (nonImageAttachments.Count == 0)
            return "";

        var items = new StringBuilder();
        foreach (var att in nonImageAttachments)
            items.Append($"<li>{WebUtility.HtmlEncode(att.FileName)}</li>");

        return $"""
            <hr>
            <h2>Attachments</h2>
            <ul>{items}</ul>
            <p><em>Files saved to: {WebUtility.HtmlEncode(folderName)}/</em></p>
            """;
    }

    private static string ProcessHtmlImages(string html, List<DownloadedAttachment> attachments)
    {
        foreach (var att in attachments.Where(
            a => a.MediaType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)))
        {
            var base64 = Convert.ToBase64String(att.Data);
            var dataUri = $"data:{att.MediaType};base64,{base64}";

            var escaped = Regex.Escape(att.FileName);
            var pattern = $@"src=""[^""]*?{escaped}[^""]*""";
            html = Regex.Replace(html, pattern, $@"src=""{dataUri}""", RegexOptions.IgnoreCase);

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
            Path.GetFileNameWithoutExtension(docPath) + "_attachments");

        Directory.CreateDirectory(dir);

        foreach (var att in attachments)
        {
            var filePath = Path.Combine(dir, att.FileName);
            File.WriteAllBytes(filePath, att.Data);
        }
    }
}
