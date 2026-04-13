namespace GrabConf;

public sealed record CommandLineOptions(
    string BaseUrl,
    string SpaceKey,
    string? User,
    string Token,
    string OutputDir,
    int MaxRequestsPerSecond,
    string? ManifestPath,
    bool Verbose);

public sealed record PageInfo(string Id, string Title, IReadOnlyList<string> AncestorTitles);

public sealed record AttachmentInfo(string FileName, string MediaType, string DownloadPath);

public sealed record DownloadedAttachment(string FileName, string MediaType, byte[] Data);

public sealed record PageMetadata(
    string? CreatorName,
    IReadOnlyList<string> Contributors,
    DateTimeOffset? CreatedDate,
    int VersionNumber,
    DateTimeOffset? LastUpdatedDate,
    long? ViewCount);

public sealed record PageContent(string Html, PageMetadata Metadata);
