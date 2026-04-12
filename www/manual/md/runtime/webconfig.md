# web.config Support

## Overview

AxonASP HTTP server reads a single `web.config` file placed at the **root** of the configured web application directory and applies URL rewrite rules, HTTP redirects, and custom error page mappings. The file follows a subset of the IIS XML schema under `<system.webServer>`.

Only the web root `web.config` is read. Files placed in subdirectories are ignored.

## Prerequisites

Enable web.config processing in `config/axonasp.toml`:

```toml
[server]
enable_webconfig = true
```

When `enable_webconfig` is `false`, the server ignores all `web.config` files and uses only the settings from `axonasp.toml`.

## Supported Sections

### URL Rewrite Rules

Declare rules inside `<system.webServer><rewrite><rules>`. Each rule is evaluated in document order against the incoming request path.

```xml
<configuration>
  <system.webServer>
    <rewrite>
      <rules>
        <rule name="Front Controller" stopProcessing="true">
          <match url="^(?!index\.asp$)(.*)" ignoreCase="true" />
          <conditions logicalGrouping="MatchAll">
            <add input="{REQUEST_FILENAME}" matchType="IsFile" negate="true" />
            <add input="{REQUEST_FILENAME}" matchType="IsDirectory" negate="true" />
          </conditions>
          <action type="Rewrite" url="/index.asp?route={R:1}" appendQueryString="true" />
        </rule>
      </rules>
    </rewrite>
  </system.webServer>
</configuration>
```

**Rule element attributes:**

| Attribute | Values | Default | Description |
|-----------|--------|---------|-------------|
| name | string | — | Display name for the rule |
| stopProcessing | true / false | false | Stop evaluating further rules when this rule matches |

**Match element attributes:**

| Attribute | Values | Default | Description |
|-----------|--------|---------|-------------|
| url | regex | — | URL pattern matched against the request path without the leading slash |
| ignoreCase | true / false | true | Perform case-insensitive matching |
| negate | true / false | false | Invert the match result |

**Conditions element attributes:**

| Attribute | Values | Default | Description |
|-----------|--------|---------|-------------|
| logicalGrouping | MatchAll / MatchAny | MatchAll | Apply AND or OR logic to all conditions in the block |

**Condition add element attributes:**

| Attribute | Values | Default | Description |
|-----------|--------|---------|-------------|
| input | see below | — | Server variable or expression to evaluate |
| matchType | IsFile / IsDirectory / Pattern | Pattern | How the input is evaluated |
| pattern | regex | — | Regular expression pattern when matchType is Pattern |
| ignoreCase | true / false | true | Case-insensitive regex matching |
| negate | true / false | false | Invert the condition result |

**Supported condition input variables:**

| Variable | Description |
|----------|-------------|
| {REQUEST_FILENAME} | Absolute filesystem path resolved from the web root for the requested URL |
| {URL} | URL path of the current request |
| {REQUEST_URI} | Same as {URL} |

**Action element attributes:**

| Attribute | Values | Default | Description |
|-----------|--------|---------|-------------|
| type | Rewrite / Redirect | — | Action to take when the rule matches |
| url | string | — | Target path or URL. Supports back-references {R:0}, {R:1}, and so on |
| appendQueryString | true / false | true | Append the original query string to the rewritten or redirected URL |
| redirectType | Permanent / Found / Temporary | Found | HTTP status code for Redirect actions (Permanent = 301, Found = 302, Temporary = 302) |

Back-references from the match pattern are available in the action URL as `{R:0}` (full match) and `{R:1}`, `{R:2}`, ... (capture groups).

---

### HTTP Redirect

Redirect all requests from the site to a fixed destination URL:

```xml
<configuration>
  <system.webServer>
    <httpRedirect enabled="true"
                  destination="https://www.newsite.com"
                  exactDestination="false"
                  childOnly="false"
                  httpResponseStatus="Permanent" />
  </system.webServer>
</configuration>
```

**httpRedirect attributes:**

| Attribute | Values | Default | Description |
|-----------|--------|---------|-------------|
| enabled | true / false | false | Activate the redirect |
| destination | URL string | — | Target redirection URL |
| exactDestination | true / false | false | Redirect every request to the exact destination URL without appending the original path |
| childOnly | true / false | false | Only redirect child paths, not the root itself |
| httpResponseStatus | Permanent / Found / Temporary | Permanent | Permanent = 301, Found = 302, Temporary = 302 |

---

### Custom Error Pages

Map HTTP status codes to custom error files:

```xml
<configuration>
  <system.webServer>
    <httpErrors>
      <error statusCode="404" path="/error-pages/404.asp" responseMode="ExecuteURL" />
      <error statusCode="500" path="/error-pages/500.html" responseMode="File" />
    </httpErrors>
  </system.webServer>
</configuration>
```

**error element attributes:**

| Attribute | Values | Description |
|-----------|--------|-------------|
| statusCode | integer | HTTP status code to intercept |
| path | string | Relative URL or file path for the custom error page |
| responseMode | File / ExecuteURL | File serves a static file; ExecuteURL executes the path as an ASP page |

Custom error mappings defined in `web.config` override the default error pages directory set via `default_error_pages_directory` in `axonasp.toml`.

## Remarks

- web.config is read only from the web root directory. Files in subdirectories are not processed.
- The `httpRedirect` directive is evaluated before rewrite rules.
- URL rewrite rules are applied before ASP script execution and before default page resolution.
- Rules are evaluated in document order. `stopProcessing="true"` prevents further rule evaluation once a rule matches.
- The `{REQUEST_FILENAME}` input resolves to an absolute filesystem path rooted at the configured web root.
- Regex patterns use Go RE2 syntax, which does not support lookahead or backreferences.
- Paths in `httpErrors` are resolved relative to the web root.

## Code Example

Route all non-file, non-directory requests through a single ASP front controller:

```xml
<configuration>
  <system.webServer>
    <rewrite>
      <rules>
        <rule name="Single Entry Point" stopProcessing="true">
          <match url="^(.+)$" ignoreCase="true" />
          <conditions logicalGrouping="MatchAll">
            <add input="{REQUEST_FILENAME}" matchType="IsFile" negate="true" />
            <add input="{REQUEST_FILENAME}" matchType="IsDirectory" negate="true" />
          </conditions>
          <action type="Rewrite" url="/index.asp" appendQueryString="true" />
        </rule>
      </rules>
    </rewrite>
  </system.webServer>
</configuration>
```
