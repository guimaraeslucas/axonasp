# System Pages and Error Pages

## Overview

AxonASP reserves two directories for internal server assets: `www/axonasp-pages/` contains styling, images, and the directory-listing template, while `www/error-pages/` contains custom HTTP error pages. Both directories are blocked from direct browser access by default.

## axonasp-pages/

**Path:** `www/axonasp-pages/`

This directory holds static assets used internally by AxonASP server features. It is not intended for application code.

**Contents:**

| Path | Description |
|------|-------------|
| `axonasp-pages/css/` | AxonASP built-in stylesheet used by the manual and system pages. |
| `axonasp-pages/images/` | Logo and icon assets for AxonASP system pages. |
| `axonasp-pages/directory-listing.html` | Go template used to render directory listings. |

### Directory Listing Template

When `enable_directory_listing = true` in `config/axonasp.toml`, the server renders a directory listing using the HTML template at `directory_listing_template`. The default template references `www/axonasp-pages/directory-listing.html`.

```toml
[server]
enable_directory_listing = true
directory_listing_template = "./www/axonasp-pages/directory-listing.html"
```

The template is a standard Go `html/template` file. It receives the following data context:

| Variable | Type | Description |
|----------|------|-------------|
| `.Title` | string | Page title, typically the directory path. |
| `.RequestPath` | string | The URL path of the requested directory. |
| `.ParentPath` | string | URL path to the parent directory. Empty for the web root. |
| `.LogoDataURI` | string | Base64-encoded data URI for the AxonASP logo image. |
| `.InlineCSS` | string | Full CSS content inlined into the `<style>` tag. |
| `.Entries` | slice | List of directory entries (files and subdirectories). |

Each `.Entries` item exposes:

| Field | Type | Description |
|-------|------|-------------|
| `.Name` | string | File or directory name. |
| `.Href` | string | URL link for the entry. |
| `.TypeLabel` | string | `File` or `Directory`. |
| `.MimeType` | string | Detected MIME type for files; `-` for directories. |
| `.SizeDisplay` | string | Human-readable file size; `-` for directories. |
| `.Modified` | string | Last-modified timestamp as a formatted string. |

**Customization:** You can replace the default template with your own by pointing `directory_listing_template` to a different file path. The template must use standard Go `html/template` syntax and the variable names listed above.

### Blocking Direct Access

Both `www/axonasp-pages/` and `www/error-pages/` are listed in `blocked_dirs` in `config/axonasp.toml`:

```toml
blocked_dirs = [
  "./www/error-pages",
  "./www/axonasp-pages",
]
```

This setting prevents clients from accessing these directories directly via HTTP. Requests to paths under these directories return a 404 response.

---

## error-pages/

**Path:** `www/error-pages/`

This directory stores custom HTTP error pages served by the AxonASP HTTP server. When a request results in an HTTP error code, the server checks this directory for a matching error page before using its built-in error response.

### Naming Convention

Error pages are matched by filename:

| Filename pattern | Description |
|-----------------|-------------|
| `<status_code>.asp` | An ASP script executed dynamically for the given status code. |
| `<status_code>.html` | A static HTML file served for the given status code. |

**Examples:**
- `403.html` — Served on HTTP 403 Forbidden responses.
- `404.html` — Served on HTTP 404 Not Found responses.
- `500.html` — Served on HTTP 500 Internal Server Error responses.
- `500.asp` — An ASP page executed and served on HTTP 500 responses.

When both an `.asp` and an `.html` file exist for the same status code, the `.asp` file takes precedence.

### Default Error Pages

AxonASP ships with default static error pages for the most common status codes (403, 404, 500). You can replace these files with your own pages by editing the files in `www/error-pages/`.

### Using web.config for Custom Error Mapping

You can also map status codes to custom pages using the `<httpErrors>` section in `web.config`. This approach allows you to use pages outside the `error-pages/` directory and to specify whether the page is executed dynamically or served as a static file:

```xml
<configuration>
  <system.webServer>
    <httpErrors>
      <error statusCode="404" path="/error-pages/404.html" responseMode="File" />
      <error statusCode="500" path="/error-pages/500.asp" responseMode="ExecuteURL" />
    </httpErrors>
  </system.webServer>
</configuration>
```

## Remarks

- Direct HTTP access to both `www/axonasp-pages/` and `www/error-pages/` is blocked by default. Do not remove these directories from `blocked_dirs` in `axonasp.toml`.
- The directory-listing template is loaded from disk on each request when directory listing is triggered. CSS and logo assets are inlined into the HTTP response; they are not served as separate requests.
- Error pages placed in `www/error-pages/` have access to all standard ASP intrinsic objects (`Request`, `Response`, `Server`, `Session`, `Application`) when using `.asp` format.
- The `axonasp-pages/css/` stylesheet is also used by the built-in documentation manual. Modifications to the stylesheet affect both the manual appearance and any custom pages that reference it.
