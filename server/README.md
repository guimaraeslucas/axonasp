# AxonASP HTTP Server (`server`)

The `server` package provides the primary standalone HTTP web server and reverse proxy for AxonASP.

## Features

- **Built-in Web Server**: Serves Classic ASP applications and static assets over HTTP/HTTPS with high throughput.
- **Request Routing & Proxying**: Handles path resolution, reverse proxying, virtual directory mapping, and MIME types.
- **ASP Pipeline Integration**: Routes incoming HTTP requests directly through `axonvm` and ASP intrinsic objects (`Request`, `Response`, `Session`, `Application`).

## Target Binary

Compiles into `./axonasp-http.exe`.
