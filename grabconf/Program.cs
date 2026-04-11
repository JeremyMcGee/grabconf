using GrabConf;

if (args.Length == 0 || args.Contains("--help") || args.Contains("-h"))
{
    PrintUsage();
    return args.Length == 0 ? 1 : 0;
}

var options = ParseArgs(args);
if (options is null)
    return 1;

using var client = new ConfluenceClient(
    options.BaseUrl, options.User, options.Token, options.MaxRequestsPerSecond);

var exporter = new WordExporter();
var tracker = options.ManifestPath is not null
    ? new ExternalSiteTracker(options.BaseUrl)
    : null;

Console.WriteLine($"Connecting to {options.BaseUrl}, space '{options.SpaceKey}'...");

var spaceName = await client.GetSpaceNameAsync(options.SpaceKey);
Console.WriteLine($"Space: {spaceName}");

var pages = await client.GetAllPagesAsync(options.SpaceKey);
Console.WriteLine($"Found {pages.Count} page(s).");

Directory.CreateDirectory(options.OutputDir);

var exported = 0;
foreach (var page in pages)
{
    Console.WriteLine($"[{++exported}/{pages.Count}] {page.Title}");
    try
    {
        var content = await client.GetPageContentAsync(page.Id);
        tracker?.Scan(content);
        var attachments = await client.GetAttachmentsAsync(page.Id);

        var downloaded = new List<DownloadedAttachment>();
        foreach (var att in attachments)
        {
            Console.WriteLine($"  Downloading attachment: {att.FileName}");
            var data = await client.DownloadAttachmentAsync(att.DownloadPath);
            downloaded.Add(new DownloadedAttachment(att.FileName, att.MediaType, data));
        }

        var pageDir = BuildPageDirectory(options.OutputDir, page.AncestorTitles);
        Directory.CreateDirectory(pageDir);

        var safeTitle = SanitizeFileName(page.Title);
        var docPath = Path.Combine(pageDir, $"{safeTitle}.docx");
        exporter.Export(docPath, page.Title, spaceName, content, downloaded);
        Console.WriteLine($"  Saved: {docPath}");
    }
    catch (HttpRequestException ex)
    {
        Console.Error.WriteLine($"  HTTP error: {ex.StatusCode} - {ex.Message}");
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"  Error: {ex.Message}");
    }
}

Console.WriteLine($"Export complete. {exported} page(s) processed.");

if (tracker is not null && options.ManifestPath is not null)
{
    tracker.WriteManifest(options.ManifestPath, options.BaseUrl, options.SpaceKey);
    Console.WriteLine($"Manifest: {options.ManifestPath}");
}

return 0;

static string SanitizeFileName(string name)
{
    var invalid = Path.GetInvalidFileNameChars();
    var sanitized = string.Concat(name.Select(c => invalid.Contains(c) ? '_' : c));
    return sanitized.Length > 200 ? sanitized[..200] : sanitized;
}

static string BuildPageDirectory(string outputDir, IReadOnlyList<string> ancestorTitles)
{
    var path = outputDir;
    foreach (var ancestor in ancestorTitles)
        path = Path.Combine(path, SanitizeFileName(ancestor));
    return path;
}

static void PrintUsage()
{
    Console.WriteLine("""
        grabconf - Export a Confluence space to Word documents

        Usage:
          grabconf --url <base-url> --space <space-key> --token <api-token> [options]

        Required:
          --url       Confluence base URL (e.g. https://mysite.atlassian.net/wiki)
          --space     Space key to export
          --token     API token or Personal Access Token

        Options:
          --user      Username/email for Basic Auth (Confluence Cloud).
                      Omit for Bearer token auth (Server/Data Center PAT).
          --output    Output directory (default: ./output)
          --rate      Max requests per second (default: 5)
          --manifest  Write a plain-text manifest of external Confluence site
                      references to the given path (e.g. ./output/manifest.txt)
          --help, -h  Show this help
        """);
}

static CommandLineOptions? ParseArgs(string[] args)
{
    string? url = null, space = null, token = null, user = null, manifest = null;
    var output = "./output";
    var rate = 5;

    for (var i = 0; i < args.Length; i++)
    {
        switch (args[i])
        {
            case "--url" when i + 1 < args.Length:
                url = args[++i];
                break;
            case "--space" when i + 1 < args.Length:
                space = args[++i];
                break;
            case "--token" when i + 1 < args.Length:
                token = args[++i];
                break;
            case "--user" when i + 1 < args.Length:
                user = args[++i];
                break;
            case "--output" when i + 1 < args.Length:
                output = args[++i];
                break;
            case "--manifest" when i + 1 < args.Length:
                manifest = args[++i];
                break;
            case "--rate" when i + 1 < args.Length:
                if (!int.TryParse(args[++i], out rate) || rate < 1)
                {
                    Console.Error.WriteLine("Error: --rate must be a positive integer.");
                    return null;
                }
                break;
            default:
                Console.Error.WriteLine($"Unknown option: {args[i]}");
                PrintUsage();
                return null;
        }
    }

    if (string.IsNullOrEmpty(url) || string.IsNullOrEmpty(space) || string.IsNullOrEmpty(token))
    {
        Console.Error.WriteLine("Error: --url, --space, and --token are required.");
        PrintUsage();
        return null;
    }

    return new CommandLineOptions(url, space, user, token, output, rate, manifest);
}
