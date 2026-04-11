using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;

namespace GrabConf;

public sealed class ConfluenceClient : IDisposable
{
    private readonly HttpClient _http;
    private readonly SemaphoreSlim _throttle = new(1, 1);
    private readonly TimeSpan _minDelay;
    private DateTime _lastRequest = DateTime.MinValue;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public ConfluenceClient(string baseUrl, string? user, string token, int maxRequestsPerSecond)
    {
        _minDelay = TimeSpan.FromMilliseconds(1000.0 / maxRequestsPerSecond);

        _http = new HttpClient
        {
            BaseAddress = new Uri(baseUrl.TrimEnd('/') + "/")
        };

        if (!string.IsNullOrEmpty(user))
        {
            var credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{user}:{token}"));
            _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", credentials);
            Log.Debug("Using Basic Auth.");
        }
        else
        {
            _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            Log.Debug("Using Bearer Token auth.");
        }
    }

    public async Task<List<PageInfo>> GetAllPagesAsync(string spaceKey, CancellationToken ct = default)
    {
        var pages = new List<PageInfo>();
        var start = 0;
        const int limit = 25;

        while (true)
        {
            var url = $"rest/api/content?spaceKey={Uri.EscapeDataString(spaceKey)}&type=page&expand=ancestors&start={start}&limit={limit}";
            Log.Debug($"Fetching pages: start={start}, limit={limit}");
            var json = await GetJsonAsync(url, ct);

            var results = json.GetProperty("results");
            if (results.GetArrayLength() == 0)
            {
                Log.Debug("No more pages to fetch.");
                break;
            }

            Log.Debug($"Received {results.GetArrayLength()} page(s) in this batch.");

            foreach (var page in results.EnumerateArray())
            {
                var id = page.GetProperty("id").GetString() ?? "";
                var title = page.GetProperty("title").GetString() ?? "";

                var ancestorTitles = new List<string>();
                if (page.TryGetProperty("ancestors", out var ancestors))
                {
                    foreach (var ancestor in ancestors.EnumerateArray())
                        ancestorTitles.Add(ancestor.GetProperty("title").GetString() ?? "");
                }

                pages.Add(new PageInfo(id, title, ancestorTitles));
            }

            if (results.GetArrayLength() < limit)
                break;

            start += results.GetArrayLength();
        }

        return pages;
    }

    public async Task<string> GetSpaceNameAsync(string spaceKey, CancellationToken ct = default)
    {
        Log.Debug($"Fetching space info for key '{spaceKey}'...");
        var url = $"rest/api/space/{Uri.EscapeDataString(spaceKey)}";
        var json = await GetJsonAsync(url, ct);
        return json.GetProperty("name").GetString() ?? spaceKey;
    }

    public async Task<string> GetPageContentAsync(string pageId, CancellationToken ct = default)
    {
        var url = $"rest/api/content/{Uri.EscapeDataString(pageId)}?expand=body.export_view";
        var json = await GetJsonAsync(url, ct);

        return json.GetProperty("body")
                    .GetProperty("export_view")
                    .GetProperty("value")
                    .GetString() ?? "";
    }

    public async Task<List<AttachmentInfo>> GetAttachmentsAsync(string pageId, CancellationToken ct = default)
    {
        var attachments = new List<AttachmentInfo>();
        var start = 0;
        const int limit = 25;

        while (true)
        {
            var url = $"rest/api/content/{Uri.EscapeDataString(pageId)}/child/attachment?start={start}&limit={limit}";
            Log.Debug($"Fetching attachments: start={start}, limit={limit}");
            var json = await GetJsonAsync(url, ct);

            var results = json.GetProperty("results");
            if (results.GetArrayLength() == 0)
                break;

            foreach (var att in results.EnumerateArray())
            {
                var title = att.GetProperty("title").GetString() ?? "unknown";

                var mediaType = "application/octet-stream";
                if (att.TryGetProperty("metadata", out var meta) &&
                    meta.TryGetProperty("mediaType", out var mt))
                {
                    mediaType = mt.GetString() ?? mediaType;
                }

                var download = att.GetProperty("_links")
                                  .GetProperty("download")
                                  .GetString() ?? "";

                attachments.Add(new AttachmentInfo(title, mediaType, download));
            }

            if (results.GetArrayLength() < limit)
                break;

            start += results.GetArrayLength();
        }

        return attachments;
    }

    public async Task<byte[]> DownloadAttachmentAsync(string downloadPath, CancellationToken ct = default)
    {
        await ThrottleAsync(ct);
        var uri = new Uri(_http.BaseAddress!, downloadPath);
        Log.Debug($"GET {uri}");
        var response = await _http.GetAsync(uri, ct);
        response.EnsureSuccessStatusCode();
        return await response.Content.ReadAsByteArrayAsync(ct);
    }

    public void Dispose()
    {
        _http.Dispose();
        _throttle.Dispose();
    }

    private async Task<JsonElement> GetJsonAsync(string relativeUrl, CancellationToken ct)
    {
        await ThrottleAsync(ct);
        Log.Debug($"GET {relativeUrl}");
        var response = await _http.GetAsync(relativeUrl, ct);
        response.EnsureSuccessStatusCode();
        return await response.Content.ReadFromJsonAsync<JsonElement>(JsonOptions, ct);
    }

    private async Task ThrottleAsync(CancellationToken ct)
    {
        await _throttle.WaitAsync(ct);
        try
        {
            var remaining = _minDelay - (DateTime.UtcNow - _lastRequest);
            if (remaining > TimeSpan.Zero)
            {
                Log.Debug($"Throttling: waiting {remaining.TotalMilliseconds:F0}ms");
                await Task.Delay(remaining, ct);
            }
            _lastRequest = DateTime.UtcNow;
        }
        finally
        {
            _throttle.Release();
        }
    }
}
