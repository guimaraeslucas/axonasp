# Install and Run AxonASP with Docker

## Overview
This page explains how to install, configure, and run official Docker container distributions of G3Pix AxonASP.

AxonASP provides two official multi-architecture Docker container variants hosted on GitHub Container Registry (`ghcr.io`):

1. **Standalone Server Container (`ghcr.io/guimaraeslucas/axonasp:latest`)**: Includes the native AxonASP HTTP engine (`axonasp-http`), FastCGI server (`axonasp-fastcgi`), CLI (`axonasp-cli`), and Model Context Protocol server (`axonasp-mcp`).
2. **Caddy Integrated Container (`ghcr.io/guimaraeslucas/axonasp:caddy`)**: Includes the Caddy Web Server with the native AxonASP Go module compiled directly in. Provides zero-process execution, automatic TLS certificate management, and Caddyfile routing.

---

## Container Image Tags

| Variant | Tag Pattern | Description |
|---|---|---|
| **Standalone Server** | `ghcr.io/guimaraeslucas/axonasp:latest` | Latest build of standard AxonASP server |
| **Standalone Server** | `ghcr.io/guimaraeslucas/axonasp:v2.3.0` | Specific semver release of standard server |
| **Standalone Server** | `ghcr.io/guimaraeslucas/axonasp:sha-xxxxxxx` | Specific commit build |
| **Caddy Edition** | `ghcr.io/guimaraeslucas/axonasp:caddy` | Latest build of Caddy-integrated edition |
| **Caddy Edition** | `ghcr.io/guimaraeslucas/axonasp:caddy-v2.3.0` | Specific semver release of Caddy edition |
| **Caddy Edition** | `ghcr.io/guimaraeslucas/axonasp:caddy-sha-xxxxxxx` | Specific commit build |

---

## 1. Standalone AxonASP Container

### Pulling and Running

Pull the standalone image:

```bash
docker pull ghcr.io/guimaraeslucas/axonasp:latest
```

Run container with default HTTP server port `8801`:

```bash
docker run -d \
  --name axonasp \
  -p 8801:8801 \
  -v /path/to/your/www:/opt/axonasp/www \
  --restart unless-stopped \
  ghcr.io/guimaraeslucas/axonasp:latest
```

### Running FastCGI or MCP Mode

The standalone container includes all AxonASP executables. You can override the default command to run FastCGI or MCP:

* **FastCGI Server (Port 9000):**
  ```bash
  docker run -d \
    --name axonasp-fastcgi \
    -p 9000:9000 \
    -v /path/to/your/www:/opt/axonasp/www \
    --restart unless-stopped \
    ghcr.io/guimaraeslucas/axonasp:latest ./axonasp-fastcgi
  ```

* **MCP Server (Port 8000):**
  ```bash
  docker run -d \
    --name axonasp-mcp \
    -p 8000:8000 \
    -v /path/to/your/www:/opt/axonasp/www \
    --restart unless-stopped \
    ghcr.io/guimaraeslucas/axonasp:latest ./axonasp-mcp
  ```

---

## 2. AxonASP Caddy Edition Container

### Pulling and Running

Pull the Caddy edition image:

```bash
docker pull ghcr.io/guimaraeslucas/axonasp:caddy
```

Run container with standard default configuration (port `8801`):

```bash
docker run -d \
  --name axonasp-caddy \
  -p 8801:8801 \
  -v /path/to/your/www:/opt/axonasp/www \
  --restart unless-stopped \
  ghcr.io/guimaraeslucas/axonasp:caddy
```

### Custom Caddyfile Configuration (Ports 80 / 443 with Automatic HTTPS)

To run Caddy as your primary web server handling SSL/TLS termination automatically, mount a custom `Caddyfile`:

```bash
docker run -d \
  --name axonasp-caddy \
  -p 80:80 \
  -p 443:443 \
  -v /path/to/your/www:/opt/axonasp/www \
  -v /path/to/Caddyfile:/opt/axonasp/Caddyfile \
  --restart unless-stopped \
  ghcr.io/guimaraeslucas/axonasp:caddy
```

Sample `Caddyfile`:

```caddyfile
example.com {
    root * /opt/axonasp/www
    route {
        axonasp
        file_server
    }
}
```

---

## Parameters and Arguments

Common Docker command arguments used in AxonASP deployment:

- `--name`:
  - Purpose: Assigns a fixed container name (e.g. `axonasp` or `axonasp-caddy`).
- `-p host_port:container_port`:
  - Purpose: Maps container network ports to host system ports (`8801` for HTTP, `9000` for FastCGI, `8000` for MCP, `80`/`443` for Caddy HTTPS).
- `-d`:
  - Purpose: Runs the container in background (detached) mode.
- `-v host_path:container_path`:
  - Purpose: Persists and customizes `www` application files (`/opt/axonasp/www`), `config` directory (`/opt/axonasp/config`), or custom `Caddyfile` (`/opt/axonasp/Caddyfile`).
- `--restart unless-stopped`:
  - Purpose: Automatically restarts container if it crashes or after system reboot.

---

## Docker Compose Setup

Below is a complete `docker-compose.yml` demonstrating how to run both containers in a local or production setup:

```yaml
services:
  # Standalone AxonASP Server
  axonasp-server:
    image: ghcr.io/guimaraeslucas/axonasp:latest
    container_name: axonasp-server
    ports:
      - "8801:8801"
      - "9000:9000"
    volumes:
      - ./www:/opt/axonasp/www
      - ./config:/opt/axonasp/config
    restart: unless-stopped

  # AxonASP Caddy Edition
  axonasp-caddy:
    image: ghcr.io/guimaraeslucas/axonasp:caddy
    container_name: axonasp-caddy
    ports:
      - "80:80"
      - "443:443"
    volumes:
      - ./www:/opt/axonasp/www
      - ./Caddyfile:/opt/axonasp/Caddyfile
    restart: unless-stopped
```

---

## Remarks

- **Mount Points:** Mount your ASP web application root directory to `/opt/axonasp/www`.
- **Permissions:** Containers run under an unprivileged `axonasp` system user with dynamic `su-exec` privilege dropping for security.
- **Health Checks:** Built-in Docker health checks verify HTTP availability every 30 seconds.
- **Cache Persistence:** You can optionally mount `/opt/axonasp/temp` to preserve compiled bytecode cache across container recreations.