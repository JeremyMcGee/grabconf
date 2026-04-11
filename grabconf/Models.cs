namespace GrabConf;

public sealed record CommandLineOptions(
    string BaseUrl,
    string SpaceKey,
    string? User,
    string Token,
    string OutputDir,
    int MaxRequestsPerSecond,
    string? ManifestPath);

public sealed record PageInfo(string Id, string Title, IReadOnlyList<string> AncestorTitles);

public sealed record AttachmentInfo(string FileName, string MediaType, string DownloadPath);

public sealed record DownloadedAttachment(string FileName, string MediaType, byte[] Data);
