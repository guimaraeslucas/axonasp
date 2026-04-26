# Write Diagnostic Output with the Global Console Object

## Overview
This page documents the global `console` object in G3Pix AxonASP. The object is available in both VBScript and JScript pages and provides four methods: `log`, `info`, `warn`, and `error`.

## Syntax
```asp
<%
' VBScript
console.log "message"
console.info "message"
console.warn "message"
console.error "message"
%>

<%@ Language=JScript %>
<%
// JScript
console.log("message");
console.info("message");
console.warn("message");
console.error("message");
%>
```

## Parameters and Arguments
- **value** (required): first argument to the console method.
- **Supported input forms:**
  - String: printed directly.
  - VBScript array: serialized to JSON and printed.
  - JScript array/object: serialized to JSON and printed.
- **Method names:**
  - `console.info(value)`
  - `console.log(value)`
  - `console.warn(value)`
  - `console.error(value)`

## Return Values
- `console.info` returns no value to the ASP page output.
- `console.log` returns no value to the ASP page output.
- `console.warn` returns no value to the ASP page output.
- `console.error` returns no value to the ASP page output.

## Remarks
- Every console output line includes date and time.
- Stream routing:
  - `console.info` writes to standard output.
  - `console.log` writes to standard output.
  - `console.warn` writes to standard error.
  - `console.error` writes to standard error.
- Console symbols in stream output:
  - `console.info`: `ℹ`
  - `console.log`: `⌨`
  - `console.warn`: `⚠`
  - `console.error`: `✖`
- File logging is controlled by `global.enable_log_files` in `config/axonasp.toml`.
- When enabled:
  - `console.log` and `console.info` are appended to `./temp/console.log`.
  - `console.warn` and `console.error` are appended to `./temp/error.log`.
- File entries do not include decorative symbols. Files store timestamp, level, and message text.

## Code Example
```asp
<%
Dim items(2)
items(0) = "alpha"
items(1) = "beta"
items(2) = 3

console.info "Starting ASP page execution"
console.log items
console.warn "Using fallback dataset"
console.error "Sample error line for diagnostics"

Response.Write "Console sample completed."
%>
```
