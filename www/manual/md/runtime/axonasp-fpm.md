# Run AxonASP-FPM for Isolated FastCGI Pools

## Overview

**G3Pix AxonASP-FPM** is the managed FastCGI process manager for multi-application deployments.
It supervises independent `axonasp-fastcgi` worker processes, applies per-pool execution identities, and restarts failed workers automatically.

Use this mode when you host multiple applications, shared hosting tenants, or any environment where one application must not affect the others.

## Requirements

- **Operating system:** Linux or macOS.
- **Privileges:** Start `axonasp-fpm` as **root**.
- **Pool files directory:** `/opt/axonasp/fpm/fpm.d/`.
- **FastCGI worker binary path:** `/opt/axonasp/axonasp-fastcgi`.

`axonasp-fpm` must start as root so it can:

- Spawn workers with per-pool UID and GID.
- Create and own socket and temp directories.
- Apply process memory controls where the operating system supports them.
- Enforce pool isolation boundaries.

## Architecture and Worker Lifecycle

`axonasp-fpm` uses a manager-plus-workers model:

1. The manager scans `/opt/axonasp/fpm/fpm.d/` for `*.conf` files.
2. For each new pool file, it starts one supervisor goroutine.
3. The supervisor validates the pool, prepares directories, and starts `axonasp-fastcgi`.
4. The manager drops worker privileges to the configured `uid` and `gid`.
5. If a worker exits unexpectedly, the supervisor restarts it after a 2-second delay.
6. Restart attempts stop when `max_restarts` is reached (unless the value is `0`).

Signal behavior:

- `SIGINT` or `SIGTERM`: Stop all pools and shut down the manager and the fastcgi workers.
- `SIGUSR2` (systemd `reload` action): Rescan pool directory, detect changed `.conf` files, gracefully stop only affected pools, and start them again with updated settings.
- `SIGHUP`: Trigger the same rescan and selective pool reload behavior.

Important operational detail:

- During selective reload, unmodified pools remain active and uninterrupted.
- Worker shutdown during reload uses graceful `SIGTERM` first and escalates to force kill only if the process does not exit within the grace window.

## Pool Configuration Directives

Each pool file in `/opt/axonasp/fpm/fpm.d/` is a TOML document.

| Directive | Required | Type | Description |
|---|---|---|---|
| `site_name` | Yes | String | Pool identifier used in manager logs and cgroup naming. Use a unique value per application. |
| `uid` | Yes | Integer | Numeric user ID used to run the worker process. |
| `gid` | Yes | Integer | Numeric group ID used to run the worker process. |
| `socket` | Yes | String | FastCGI endpoint passed to `--fastcgi.server_port`. Supports Unix socket paths or TCP endpoints. |
| `config_file` | Yes | String | Absolute path to the AxonASP TOML config file used by the worker (`--config.config_file`). |
| `global_asa` | No | String | Optional directory passed to the worker as `--config.global_asa`. Use this to force one explicit `global.asa` boundary per pool. |
| `app_path` | Yes | String | Worker current directory and web root. The FPM supervisor sets this as the worker's current working directory and also passes it as `--server.web_root` to the FastCGI process. This guarantees the worker serves files from the correct directory instead of falling back to the default `./www` relative path. Must be an existing directory. |
| `memory_limit_mb` | Yes | Integer | Memory ceiling in MB. The manager exports memory-related environment variables and attempts cgroup enforcement. |
| `max_restarts` | Yes | Integer | Maximum restart attempts after crashes. Use `0` for unlimited restarts. |
| `tmp_dir` | No | String | Temporary directory for the pool. Defaults to `/opt/axonasp/temp/` when omitted. |

## Supported Socket Values

`socket` accepts these endpoint styles:

- Unix socket with prefix: `unix:/var/run/axonasp/example.sock`
- TCP host and port: `127.0.0.1:9100`
- TCP port only: `9000`

If you use a Unix socket, the manager creates the parent directory, changes ownership to the pool UID and GID, and removes stale socket files before worker start.

## CLI Flags Passed to the FastCGI Worker

The FPM supervisor builds the worker command line automatically from the pool configuration file. The following flags are always passed:

| CLI Flag | Source | Description |
|---|---|---|
| `--fastcgi.server_port` | `socket` | FastCGI listen endpoint. |
| `--config.config_file` | `config_file` | Path to the AxonASP TOML configuration file. |
| `--global.temp_dir` | `tmp_dir` | Temporary directory for runtime files such as sessions and cache. |
| `--server.web_root` | `app_path` | Web root directory. Explicitly overrides `server.web_root` from the TOML config so the worker always serves from the correct application directory. |

When the corresponding pool directive is set, these optional flags are also appended:

| CLI Flag | Source | Condition |
|---|---|---|
| `--config.global_asa` | `global_asa` | Added when `global_asa` is set and not blank. |
| `--pool.name` | `site_name` | Added when `site_name` is set and not blank.

**Do not set `--server.web_root` in the worker TOML config when running under FPM.** The FPM supervisor always passes this flag from the pool `app_path` directive, and its value takes precedence over the TOML file.

## Complete Pool Example

You can use this as a production baseline and adjust values for each application or use the example file in `/opt/axonasp/fpm/fpm.d/example.conf` as a template:

```toml
site_name = "example.com"
uid = 1001
gid = 1001
socket = "/var/run/axonasp/example.com.sock"
config_file = "/opt/axonasp/config/axonasp.toml"
global_asa_path = "/opt/axonasp/sites/example.com"
app_path = "/opt/axonasp/"
memory_limit_mb = 256
max_restarts = 5
tmp_dir = "/opt/axonasp/temp/"
```

## Permissions and Isolation Checklist

Before you start `axonasp-fpm`, ensure the configured `uid` and `gid` have:

- **Read access** to `config_file`.
- **Read and execute access** to `app_path` and application files.
- **Write access** to `tmp_dir`.
- **Write access** to the Unix socket directory when using a Unix socket.

If permissions are incorrect, the pool fails to start or restarts repeatedly.

## Start and Operate the Manager

Example startup:

```bash
sudo /opt/axonasp/axonasp-fpm
```

For production, run `axonasp-fpm` as a **systemd service** instead of a manual shell process.

### Create a systemd Service (Recommended)

Create `/etc/systemd/system/axonasp-fpm.service`:

```ini
[Unit]
Description=AxonASP-FPM
After=network.target
Wants=network-online.target

[Service]
Type=simple
User=root
Group=root
ExecStart=/opt/axonasp/axonasp-fpm
Restart=on-failure
RestartSec=2
KillSignal=SIGTERM
TimeoutStopSec=30
StandardOutput=journal
StandardError=journal
SyslogIdentifier=axonasp-fpm

[Install]
WantedBy=multi-user.target
```

Enable and start the service:

```bash
sudo systemctl daemon-reload
sudo systemctl enable axonasp-fpm
sudo systemctl start axonasp-fpm
```

Check status and logs:

```bash
sudo systemctl status axonasp-fpm
sudo journalctl -u axonasp-fpm -f
```

Common operations:

```bash
# Direct process mode: add new pool file, then trigger rescan
sudo kill -HUP <fpm_manager_pid>

# Direct process mode: graceful stop
sudo kill -TERM <fpm_manager_pid>

# systemd mode: graceful stop/start, full restart and reload
sudo systemctl stop axonasp-fpm
sudo systemctl start axonasp-fpm
sudo systemctl restart axonasp-fpm
sudo systemctl reload axonasp-fpm #Reloads only modified pools; keeps unmodified pools running.

```

## IIS Administrator Translation Guide

If you are migrating from a Windows IIS model, use the mappings below.

| AxonASP Concept | IIS Equivalent | Practical Meaning |
|---|---|---|
| One pool `.conf` file in `fpm.d` | One IIS Application Pool | Each pool is an isolated worker boundary for one application context. |
| `uid` and `gid` in pool config | AppPool Identity | Isolation requires separate pool files with distinct Unix users and groups. |
| `global_asa` in pool config | Virtual directory application boundary | Defining `global_asa` pins startup events and application scope to one explicit path per pool. |
| `socket` endpoint per pool | Handler endpoint per app pool | Never share one socket across isolated applications. |

Migration guidance:

- Use one `.conf` file per isolated application, similar to one dedicated IIS Application Pool per site.
- Use different `uid`/`gid` values to replicate identity isolation between applications.
- Set `global_asa` explicitly for each pool when multiple applications exist under one server tree.
- Prefer absolute paths for `config_file`, `global_asa`, `app_path`, and `tmp_dir` in production to avoid path-resolution ambiguity.

## Memory Control Notes

For each worker, the manager exports:

- `GOMEMLIMIT=<memory_limit_mb>MiB`
- `GLOBAL_GOLANG_MEMORY_LIMIT_MB=<memory_limit_mb>MiB`
- `GLOBAL_TEMP_DIR=<tmp_dir>`
- `TMPDIR=<tmp_dir>`

On Linux, the manager also attempts cgroup v2 memory enforcement under:

`/sys/fs/cgroup/axonasp/<site_name>`

If cgroup delegation or permissions are missing, the manager logs a warning and continues running the pool.

## Standalone vs FPM Boundaries

Do not mix unmanaged standalone workers and managed pools for the same application endpoint.

- Do not start `axonasp-fastcgi` manually on a socket or port that an FPM pool already owns.
- Do not manage crash recovery scripts for workers already supervised by FPM.
- Do not share one Unix socket path across multiple pools.

Use standalone FastCGI only for single-application or explicitly manual deployments.
Use AxonASP-FPM as the default production model for multi-application hosting and when you need to use more than one global.asa context.