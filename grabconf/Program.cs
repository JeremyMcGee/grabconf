using System.Security.Cryptography;
using System.Text;
using GrabConf;
using Spectre.Console;

if (args.Length == 0 || args.Contains("--help") || args.Contains("-h"))
{
    PrintUsage();
    return args.Length == 0 ? 1 : 0;
}

var options = ParseArgs(args);
if (options is null)
    return 1;

Log.Verbose = options.Verbose;

AnsiConsole.Write(new Rule("[blue]grabconf[/]").LeftJustified());

Log.Debug($"Base URL: {options.BaseUrl}");
Log.Debug($"Space key: {options.SpaceKey}");
Log.Debug($"Auth mode: {(options.User is not null ? "Basic Auth (user: " + options.User + ")" : "Bearer Token")}");
Log.Debug($"Rate limit: {options.MaxRequestsPerSecond} req/s");
Log.Debug($"Output directory: {Path.GetFullPath(options.OutputDir)}");
if (options.ManifestPath is not null)
    Log.Debug($"Manifest path: {options.ManifestPath}");

using var client = new ConfluenceClient(
    options.BaseUrl, options.User, options.Token, options.MaxRequestsPerSecond);

var exporter = new HtmlExporter();
var tracker = options.ManifestPath is not null
    ? new ExternalSiteTracker(options.BaseUrl)
    : null;

Log.Info($"Connecting to {options.BaseUrl}, space '{options.SpaceKey}'...");

var spaceName = await client.GetSpaceNameAsync(options.SpaceKey);
Log.Success($"Space: {spaceName}");

Log.Info("Fetching page list...");
var pages = await client.GetAllPagesAsync(options.SpaceKey);
Log.Success($"Found {pages.Count} page(s).");

Directory.CreateDirectory(options.OutputDir);
Log.Debug($"Output directory created/verified: {Path.GetFullPath(options.OutputDir)}");

var exported = 0;
var pageTitles = new Dictionary<string, string>();
foreach (var page in pages)
{
    AnsiConsole.MarkupLine($"[bold][[{++exported}/{pages.Count}]][/] {page.Title.EscapeMarkup()}");
    try
    {
        Log.Debug($"Fetching content for page '{page.Title}' (id={page.Id})...");
        var pageContent = await client.GetPageContentAsync(page.Id);
        Log.Debug($"Content size: {pageContent.Html.Length:N0} characters");

        tracker?.Scan(pageContent.Html);

        Log.Debug($"Fetching view count for page '{page.Title}'...");
        var viewCount = await client.GetPageViewCountAsync(page.Id);
        var metadata = pageContent.Metadata with { ViewCount = viewCount };

        if (Log.Verbose)
        {
            Log.Debug($"Creator: {metadata.CreatorName ?? "unknown"}");
            Log.Debug($"Contributors: {(metadata.Contributors.Count > 0 ? string.Join(", ", metadata.Contributors) : "none")}");
            Log.Debug($"Created: {metadata.CreatedDate?.ToString("yyyy-MM-dd HH:mm:ss") ?? "unknown"}");
            Log.Debug($"Version: {metadata.VersionNumber}, Last updated: {metadata.LastUpdatedDate?.ToString("yyyy-MM-dd HH:mm:ss") ?? "unknown"}");
            Log.Debug($"Views: {viewCount?.ToString() ?? "N/A"}");
        }

        Log.Debug($"Fetching attachments for page '{page.Title}'...");
        var attachments = await client.GetAttachmentsAsync(page.Id);
        Log.Debug($"Found {attachments.Count} attachment(s) for '{page.Title}'");

        var downloaded = new List<DownloadedAttachment>();
        foreach (var att in attachments)
        {
            Log.Detail($"Downloading attachment: {att.FileName}");
            var data = await client.DownloadAttachmentAsync(att.DownloadPath);
            Log.Debug($"Downloaded {att.FileName} ({data.Length:N0} bytes, {att.MediaType})");
            downloaded.Add(new DownloadedAttachment(att.FileName, att.MediaType, data));
        }

        var pageDir = BuildPageDirectory(options.OutputDir, page.AncestorTitles);
        Directory.CreateDirectory(pageDir);

        var safeTitle = SanitizeFileName(page.Title);
        var docPath = Path.Combine(pageDir, $"{safeTitle}.html");
        Log.Debug($"Creating document: {docPath}");
        exporter.Export(docPath, page.Title, spaceName, pageContent.Html, downloaded, metadata);
        pageTitles[Path.GetFullPath(docPath)] = page.Title;
        Log.Success($"Saved: {docPath}");
    }
    catch (HttpRequestException ex)
    {
        Log.Error($"HTTP error: {ex.StatusCode} - {ex.Message}");
        Log.Debug($"Stack trace: {ex.StackTrace}");
    }
    catch (Exception ex)
    {
        Log.Error($"Error: {ex.Message}");
        Log.Debug($"Stack trace: {ex.StackTrace}");
    }
}

AnsiConsole.Write(new Rule().RuleStyle("green"));
Log.Success($"Export complete. {exported} page(s) processed.");

Log.Info("Generating index files...");
HtmlExporter.GenerateIndexFiles(options.OutputDir, spaceName, pageTitles);
Log.Success("Index files generated.");

if (tracker is not null && options.ManifestPath is not null)
{
    tracker.WriteManifest(options.ManifestPath, options.BaseUrl, options.SpaceKey);
    Log.Success($"Manifest written: {options.ManifestPath}");
}

return 0;

static string SanitizeName(string name, int maxLength)
{
    const int hashLength = 4;
    var prefixLength = maxLength - hashLength - 1;

    var sb = new StringBuilder(name.Length);
    foreach (var c in name)
    {
        if (char.IsLetterOrDigit(c))
            sb.Append(c);
        else if (sb.Length > 0 && sb[sb.Length - 1] != '-')
            sb.Append('-');
    }

    var sanitized = sb.ToString().Trim('-');
    if (sanitized.Length == 0)
        sanitized = "page";

    if (sanitized.Length <= maxLength)
        return sanitized;

    var hash = Convert.ToHexStringLower(
        SHA256.HashData(Encoding.UTF8.GetBytes(name)))[..hashLength];
    return $"{sanitized[..prefixLength]}_{hash}";
}

static string SanitizeFileName(string name) => SanitizeName(name, maxLength: 40);

static string SanitizeFolderName(string name) => SanitizeName(name, maxLength: 30);

static string BuildPageDirectory(string outputDir, IReadOnlyList<string> ancestorTitles)
{
    var path = outputDir;
    foreach (var ancestor in ancestorTitles)
        path = Path.Combine(path, SanitizeFolderName(ancestor));
    return path;
}

static void PrintUsage()
{
    Console.WriteLine("""
        grabconf - Export a Confluence space to HTML

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
          --verbose, -v  Enable verbose/debug logging
          --help, -h  Show this help
        """);
}

static CommandLineOptions? ParseArgs(string[] args)
{
    string? url = null, space = null, token = null, user = null, manifest = null;
    var output = "./output";
    var rate = 5;
    var verbose = false;

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
            case "--verbose" or "-v":
                verbose = true;
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

    return new CommandLineOptions(url, space, user, token, output, rate, manifest, verbose);
}
