# Installing AxonASP from Linux Packages

## Overview

AxonASP publishes pre-built binary packages for Linux on every tagged release. These packages install all server binaries, the default configuration, the web root, and the required directory structure under `/opt/axonasp`. Installation is handled by the standard package manager of your distribution.

Packages are available in three formats:

- **.deb** — Debian, Ubuntu, and derivatives
- **.rpm** — Red Hat, Fedora, CentOS, Rocky Linux, and derivatives
- **.apk** — Alpine Linux

---

## Prerequisites

- A 64-bit (amd64) Linux system
- Root or `sudo` access
- Internet access to reach the GitHub Releases page, or a locally downloaded package file

---

## Downloading a Package

All packages are published to the GitHub Releases page. The download URL follows this pattern:

```
https://github.com/guimaraeslucas/axonasp/releases/download/<version>/<file>.<extension>
```

Replace `<version>` with the release tag (for example, `v2.0.1`), `<file>` with the package filename, and `<extension>` with `deb`, `rpm`, or `apk`.

**Debian/Ubuntu example:**

```bash
wget https://github.com/guimaraeslucas/axonasp/releases/download/v2.0.1/axonasp_2.0.1_linux_amd64.deb
```

**Red Hat/Fedora/Rocky example:**

```bash
wget https://github.com/guimaraeslucas/axonasp/releases/download/v2.0.1/axonasp-2.0.1-1.x86_64.rpm
```

**Alpine example:**

```bash
wget https://github.com/guimaraeslucas/axonasp/releases/download/v2.0.1/axonasp_2.0.1_x86_64.apk
```

You can check the actual filenames for a given release on the releases page at `https://github.com/guimaraeslucas/axonasp/releases`.

---

## Installing the Package

### Debian and Ubuntu

```bash
sudo dpkg -i axonasp_2.0.1_linux_amd64.deb
```

Or, to automatically resolve any missing recommendations:

```bash
sudo apt-get install -f ./axonasp_2.0.1_linux_amd64.deb
```

### Red Hat, Fedora, CentOS, Rocky Linux

```bash
sudo rpm -i axonasp-2.0.1-1.x86_64.rpm
```

Or using `dnf`:

```bash
sudo dnf localinstall axonasp-2.0.1-1.x86_64.rpm
```

### Alpine Linux

```bash
sudo apk add --allow-untrusted axonasp_2.0.1_x86_64.apk
```

---

## What Gets Installed

After a successful installation, the following layout is created:

| Path | Description |
|---|---|
| `/opt/axonasp/` | Main installation directory |
| `/opt/axonasp/axonasp-http` | HTTP web server binary |
| `/opt/axonasp/axonasp-fastcgi` | FastCGI application server binary |
| `/opt/axonasp/axonasp-cli` | Command-line interpreter binary |
| `/opt/axonasp/axonasp-mcp` | MCP server binary |
| `/opt/axonasp/axonasp-testsuite` | Automated test suite runner binary |
| `/opt/axonasp/axonasp-service` | Service wrapper helper binary |
| `/opt/axonasp/config/axonasp.toml` | Default configuration file |
| `/opt/axonasp/www/` | Default web root and documentation |
| `/opt/axonasp/mcp/` | MCP server resource files |
| `/opt/axonasp/global.asa` | Default global.asa application lifecycle file |
| `/opt/axonasp/temp/cache/` | Bytecode script cache directory |
| `/opt/axonasp/temp/session/` | Session storage directory |

---

## Symbolic Links

The package creates symbolic links so that the server binaries are accessible from any terminal without specifying the full path.

| Symlink | Target |
|---|---|
| `/usr/bin/axonasp-http` | `/opt/axonasp/axonasp-http` |
| `/usr/bin/axonasp-cli` | `/opt/axonasp/axonasp-cli` |
| `/usr/bin/axonasp-fastcgi` | `/opt/axonasp/axonasp-fastcgi` |
| `/usr/bin/axonasp-mcp` | `/opt/axonasp/axonasp-mcp` |
| `/usr/bin/axonasp-testsuite` | `/opt/axonasp/axonasp-testsuite` |
| `/etc/axonasp/config/axonasp.toml` | `/opt/axonasp/config/axonasp.toml` |
| `/var/cache/axonasp/` | `/opt/axonasp/temp/cache/` |
| `/var/lib/axonasp/session/` | `/opt/axonasp/temp/session/` |

Because the binaries are linked under `/usr/bin`, you can start the server from any working directory:

```bash
axonasp-http
```

The configuration symlink at `/etc/axonasp/config/axonasp.toml` provides a standard location for system administrators to locate and edit the configuration file without navigating to `/opt/axonasp`.

---

## Starting the Server After Installation

After the package installs, start the HTTP server directly:

```bash
axonasp-http
```

To run it as a persistent background service, follow the instructions in the **Running AxonASP as a Linux Service** page of this manual. The package includes `install-service.sh` and `uninstall-service.sh` scripts under `/opt/axonasp/` that automate systemd unit file creation.

```bash
cd /opt/axonasp
sudo bash install-service.sh
sudo systemctl start axonasp
sudo systemctl enable axonasp
```

---

## Remarks

- All packages target the **amd64 (x86-64)** architecture. ARM builds are not included in the pre-built packages; compile from source using `build.sh` for other architectures.
- The `temp/` subdirectories are created as empty directories by the package. The server writes bytecode caches and session data there at runtime.
- The configuration file at `/opt/axonasp/config/axonasp.toml` contains default values for all settings. Edit it before starting the server to configure the web root path, port, session behavior, and other options.
- The `www/tests/` directory is excluded from the distributed package. It is only present in the source repository.
- Package builds are triggered automatically on every push to the `main` branch and on every version tag. Non-tagged builds use a commit-count-based version number (`2.0.<commit-count>.<short-hash>`). Tagged releases use the clean tag version (for example, `2.0.1`).
