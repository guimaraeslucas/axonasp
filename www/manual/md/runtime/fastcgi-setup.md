# Deploy AxonASP with FastCGI

## Overview

G3Pix AxonASP supports three FastCGI deployment models.
Choose the model based on your isolation, automation, and operational requirements.

For production environments that host multiple applications, **AxonASP-FPM is the default and recommended model**.

## Deployment Models

| Model | Process Ownership | global.asa Scope | Best Use Case | Operational Risk |
|---|---|---|---|---|
| Single Application Standalone | One `axonasp-fastcgi` process | One `global.asa` for that process | Small single-site deployments with minimal memory use | Low |
| Multiple Manual Standalone | Multiple manually started `axonasp-fastcgi` processes | One `global.asa` per process | Small systems that need custom per-process tuning | Medium to High |
| AxonASP-FPM Managed Pools | One `axonasp-fpm` manager supervising multiple workers | One isolated `global.asa` context per pool configuration | Multi-application production and shared hosting | Low |

## Model 1: Single Application Standalone

Run one FastCGI worker directly.

```bash
./axonasp-fastcgi
```

This model has the smallest footprint because only one worker process runs.
It is the correct choice when you host one application and do not require pool orchestration.

### global.asa Resolution in Standalone Mode

At startup, the worker resolves `global.asa` in this order:

1. Explicit `--config.global_asa` directory.
2. Current working directory.
3. `server.web_root` from the active TOML config (or `--server.web_root` CLI flag if set).

If `--config.global_asa` is explicitly provided and `global.asa` is missing in that directory, startup fails with an internal 500 state.

Path enforcement for Step 3:

- If `server.web_root` is absolute, FastCGI resolves it strictly as absolute.
- If `server.web_root` is relative, FastCGI resolves it relative to the current working directory.
- FastCGI does not rewrite absolute `server.web_root` values to `./www`.
- The `--server.web_root` CLI flag always takes precedence over the TOML `server.web_root` value.

## Model 2: Multiple Manual Standalone Processes

Run multiple `axonasp-fastcgi` processes manually, each with explicit arguments.

```bash
./axonasp-fastcgi --config.config_file /opt/axonasp/config/site-a.toml --config.global_asa /opt/axonasp/sites/site-a
./axonasp-fastcgi --config.config_file /opt/axonasp/config/site-b.toml --config.global_asa /opt/axonasp/sites/site-b --fastcgi.server_port unix:/var/run/axonasp/site-b.sock
```

Use this model when you need granular process-by-process control and your environment is still small enough for manual operations.

Administrative responsibility in this model:

- You create startup scripts.
- You handle daemonization.
- You monitor health and restart crashes.
- You prevent port and socket conflicts.

If you do not enforce those controls, orphan workers and configuration drift become likely.

## Model 3: AxonASP-FPM Managed Pools (Recommended)

Use `axonasp-fpm` to supervise independent FastCGI workers as isolated pools.

This model is similar to PHP-FPM or application-pool orchestration:

- Automatic worker start and restart.
- Per-pool UID and GID execution.
- Per-pool memory limits and guardrails.
- Per-pool socket and path isolation.

**Production recommendation:** Use this model for multi-application servers and shared hosting.

### Mandatory Administrative Requirement

Start `axonasp-fpm` as **root**.
The manager needs root initialization to:

- Drop worker privileges to the configured pool user and group.
- Prepare and own socket and temporary directories.
- Apply process and memory controls correctly.
- Enforce security boundaries between pools.

### Pool Reload Behavior

AxonASP-FPM reloads changed pool files with `SIGUSR2` (or `SIGHUP`) using selective restart behavior:

- New `.conf` files start new pools.
- Modified `.conf` files gracefully restart only their own pools.
- Unmodified pools continue running without interruption.
- Worker shutdown on reload sends `SIGTERM` first, then force-kill only if the worker does not exit during the grace window.

This keeps active applications stable while allowing targeted configuration rollout.

## FastCGI Worker Endpoint Formats

Standalone workers and FPM-managed workers both accept these listen endpoint styles:

- `9000`
- `127.0.0.1:9000`
- `:9000`
- `unix:/var/run/axonasp/app.sock`

When endpoint configuration is empty, FastCGI falls back to `9000`.

## Reverse Proxy Requirements

Always run FastCGI behind a reverse proxy.
The proxy must pass CGI variables required for script path resolution, especially in multi-host routing.

Minimum required parameters for reliable execution:

- `DOCUMENT_ROOT`
- `SCRIPT_NAME`
- `REQUEST_METHOD`
- `QUERY_STRING`
- `SERVER_NAME`
- `SERVER_PORT`

## Do Not Mix Standalone and Managed Ownership

Do not combine manual standalone process scripts and AxonASP-FPM for the same application endpoint.

Incorrect patterns:

- Running `axonasp-fastcgi` manually on a socket already assigned to an FPM pool.
- Creating external restart loops for workers already managed by FPM.
- Sharing one socket path across unrelated applications.

Correct pattern:

- Use standalone ownership for fully manual deployments.
- Use FPM ownership for automated multi-application deployments.
- Keep ownership boundaries explicit and non-overlapping.

## Standalone Startup Flags

`axonasp-fastcgi` supports these primary startup flags:

| Flag | Description |
|---|---|
| `-c`, `--config.config_file` | Set the AxonASP TOML file path. |
| `--fastcgi.server_port` | Override listen endpoint (port, host:port, `:port`, or `unix:/path/socket`). |
| `--config.global_asa` | Set explicit directory that must contain `global.asa`. |
| `--server.web_root` | Override the web root directory. Takes precedence over `server.web_root` from the TOML config file. When run under FPM, this flag is automatically injected from the pool `app_path` directive and should not be set in the TOML config. |
| `--global.temp_dir` | Override the temporary directory for runtime files such as sessions and cache. |
| `--pool.name` | Set a pool name identifier used in worker log output. Useful for distinguishing multiple workers in logs. |
| `-a`, `--about` | Print product and licensing information, then exit. |

**Precedence rule:** CLI flags always override the corresponding values from the TOML configuration file. If a flag is not set on the command line, the worker falls back to the TOML value or its built-in default.

Examples:

```bash
./axonasp-fastcgi --fastcgi.server_port 9001
./axonasp-fastcgi --fastcgi.server_port unix:/var/run/axonasp/site-a.sock
./axonasp-fastcgi --config.config_file /opt/axonasp/config/site-a.toml --config.global_asa /opt/axonasp/sites/site-a
```

## IIS Administrator Translation Guide

Use this mapping if your operational model comes from IIS:

| AxonASP Runtime Term | IIS Term | Migration Note |
|---|---|---|
| FPM pool `.conf` file | Application Pool | Isolate each application with a dedicated pool file. |
| `uid` and `gid` per pool | AppPool Identity | Use unique Unix identities for boundary isolation. |
| `global_asa` per pool | Application boundary in a virtual directory | Explicit `global_asa` makes startup scope deterministic per app context. |
| One Unix socket per pool | Dedicated handler endpoint | Do not share sockets between unrelated applications. |

Production recommendation:

- Use absolute paths in pool and runtime settings wherever possible (`config_file`, `global_asa`, `app_path`, `tmp_dir`, socket path).
- Keep one explicit ownership model per endpoint: standalone or FPM, never both.

## Remarks

- FastCGI execution parity is maintained with the AxonASP HTTP runtime for ASP libraries and language features.
- FastCGI does not serve static files directly. The reverse proxy serves static content.
- IIS native FastCGI is not supported. Use the IIS reverse proxy path described in the IIS runtime documentation.
