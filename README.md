# grabconf

A command-line tool that exports an entire Confluence space to Word documents (`.docx`), including page attachments.

## Features

- Exports every page in a Confluence space to an individual `.docx` file
- Embeds images from page content directly into the Word document as base64 data URIs
- Downloads and saves all page attachments to a companion folder alongside each document
- Rate-limits API requests to avoid overloading the Confluence server
- Supports both Confluence Cloud (Basic Auth) and Server/Data Center (Bearer PAT) authentication
- Continues exporting remaining pages if an individual page fails
- Optionally generates a plain-text manifest of external Confluence sites referenced across the space, with reference counts

## Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0) or later

## Build

```shell
dotnet build grabconf/grabconf.csproj
```

## Usage

```
grabconf --url <base-url> --space <space-key> --token <api-token> [options]
```

### Required arguments

| Argument  | Description                                                        |
|-----------|--------------------------------------------------------------------|
| `--url`   | Confluence base URL (e.g. `https://mysite.atlassian.net/wiki`)     |
| `--space` | Space key to export                                                |
| `--token` | API token (Cloud) or Personal Access Token (Server / Data Center)  |

### Optional arguments

| Argument   | Default    | Description                                                                                  |
|------------|------------|----------------------------------------------------------------------------------------------|
| `--user`   | *(none)*   | Username or email for Basic Auth (Confluence Cloud). Omit for Bearer token auth (Server PAT). |
| `--output` | `./output` | Directory where `.docx` files and attachment folders are written.                             |
| `--rate`     | `5`        | Maximum API requests per second.                                                             |
| `--manifest` | *(none)*   | Path for a plain-text manifest listing external Confluence sites referenced in the space.     |
| `--help`     |            | Show help text.                                                                              |

### Examples

**Confluence Cloud** (Basic Auth with email + API token):

```shell
grabconf --url https://mysite.atlassian.net/wiki \
         --space DEV \
         --user user@example.com \
         --token ATATT3x... \
         --output ./export \
         --manifest ./export/external-sites.txt
```

**Confluence Server / Data Center** (Personal Access Token):

```shell
grabconf --url https://confluence.internal.com \
         --space TEAM \
         --token NjM4OTY3... \
         --rate 3
```

## Output structure

```
output/
├── Top Level Page.docx
├── Top Level Page_attachments/
│   └── notes.txt
├── Getting Started/
│   ├── Installation.docx
│   ├── Installation_attachments/
│   │   └── setup.sh
│   └── Installation/
│       └── Linux Setup.docx
├── Architecture/
│   ├── Overview.docx
│   └── Overview/
│       ├── Backend.docx
│       └── Backend_attachments/
│           └── diagram.png
└── Another Top Level Page.docx
```

The output folder hierarchy mirrors the page tree in Confluence. Top-level pages are written directly into the output directory. Child pages are placed in a subfolder named after their parent, grandchildren in a subfolder of that, and so on. Ancestor titles are sanitized the same way as file names (invalid characters replaced with `_`, truncated to 200 characters).

Each Word document begins with a three-line context header:

| Line         | Example                     |
|--------------|-----------------------------|
| **Space**    | Development                 |
| **Title**    | Installation                |
| **Exported** | 2025-07-14 09:32:15         |

Below the header, the page content and any attachment listing follow.

- Image attachments referenced in the page body are embedded inline in the Word document.
- All attachments (including images) are also saved to a `<page>_attachments/` folder next to the document.
- Non-image attachments are listed in an **Attachments** section at the end of the document.

## External site manifest

When `--manifest <path>` is supplied, the tool scans each page's HTML for hyperlinks that point to
other Confluence instances and writes a summary file:

```
External Confluence Sites Manifest
Generated: 2025-07-14 09:32:15
Source:    https://mysite.atlassian.net/wiki (space: DEV)
---

https://partner.atlassian.net/wiki    12
https://confluence.vendor.com          3
https://docs.acme.com/wiki             1

Total: 16 reference(s) to 3 external site(s)
```

A URL is classified as a Confluence link when its path matches common Confluence patterns
(`/wiki/…`, `/display/…`, `/pages/viewpage.action…`, `/confluence/…`, or any `*.atlassian.net` host).
Links back to the source site are excluded.

## Rate limiting

Requests to the Confluence REST API are throttled to the rate specified by `--rate` (default 5 requests/second). A semaphore-based token bucket enforces a minimum delay between consecutive requests to stay within the limit.

## How it works

1. Fetches the space name from the Confluence REST API.
2. Lists all pages in the specified space via `/rest/api/content` (with `expand=ancestors`), handling pagination automatically.
3. For each page, fetches the rendered HTML content using the `body.export_view` expansion.
4. Retrieves all attachments for the page (also paginated) and downloads each file.
5. Replaces image `src` URLs in the HTML with base64 data URIs from the downloaded attachments so images render inline in Word.
6. Creates a `.docx` file using the Open XML SDK, embedding the full HTML via an `AltChunk` import part. A context header (space name, title, export date) is written at the top of each document.
7. Places the `.docx` in a folder hierarchy derived from the page's ancestor chain, mirroring the Confluence page tree.
8. Saves all attachment files to a companion folder on disk.
9. If `--manifest` was specified, writes a plain-text file listing every external Confluence site found in the exported pages together with the number of times it was referenced.

## License

This project is provided as-is with no warranty. Use at your own risk.
