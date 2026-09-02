# AxonASP FastCGI (`fastcgi`)

The `fastcgi` package implements the FastCGI server protocol for AxonASP, enabling high-performance integration with web servers such as Microsoft IIS, Nginx, Apache, and Caddy.

## Features

- **FastCGI Protocol Handler**: Efficiently handles incoming FastCGI requests, managing request streams, headers, and environment parameters.
- **ASP Execution Bridge**: Integrates directly with `axonvm` to process Classic ASP scripts with full intrinsic object and session support.
- **IIS & Nginx Ready**: Supports standard FastCGI worker execution modes and named pipes/TCP sockets.

## Target Binary

Compiles into `./axonasp-fastcgi.exe`.
