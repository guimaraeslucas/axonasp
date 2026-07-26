# AxonHTA

AxonHTA is a desktop runtime that lets you build cross-platform desktop applications using **pure VBScript + HTML + CSS** — no IIS, no JavaScript framework, no Go code required.

It is a modern reimagining of the classic Windows HTA (HTML Application): the same development experience, but with the engine upgraded from IE to modern Chromium, and the ability to run cross-platform.

## How It Works

```
VBScript (.hta / .asp) → AxonASP VM compiles and executes → embedded HTTP server → Chromium app window
```

1. `axonhta.exe` starts an internal HTTP server (listening on a random local port).
2. Your `.hta` / `.asp` files are executed on this server, exactly as they would be under IIS.
3. AxonHTA automatically finds Chrome, Edge, or Chromium on the system and launches it in **app mode** — a borderless standalone window with no address bar or tabs, giving a native desktop feel.
4. User clicks buttons, submits forms → HTTP request → VBScript processes → returns a new page (or partial update via HTMX).
5. Closing the window exits the application automatically (detected via heartbeat).

If no Chromium-based browser is found, AxonHTA falls back to the system default browser.

### Process Lifecycle

AxonHTA injects a small script into every HTML response that sends a heartbeat (`HEAD /__heartbeat__`) every 5 seconds. When the browser window is closed, heartbeats stop arriving, and the server exits after a 15-second grace period. The script also sends a final heartbeat via `navigator.sendBeacon` on `pagehide`/`beforeunload` to cover form-submission navigation gaps (where the old page's JS stops before the new page's JS starts). This works reliably regardless of whether Chrome forks its launcher process (a common Windows behavior that breaks naive `cmd.Wait()` approaches).

### Desktop Experience

The injected script also disables the browser context menu (right-click) and drag-and-drop of page elements, providing a more native desktop feel.

## Building

### Prerequisites

- Clone the AxonASP source code locally.
- Go 1.22+.
- A Chromium-based browser installed (Chrome, Edge, Chromium, Brave, etc.).

### Build

```bash
cd axonasp/axonhta
go build -o axonhta.exe
```

No CGO required — AxonHTA is pure Go. The window is provided by launching the system's Chromium-based browser in app mode.

### Windows: Hiding the Console Window

By default, a console window appears alongside the app window. To hide it:

```bash
go build -ldflags="-H windowsgui" -o axonhta.exe
```

Or use the included build script (which applies this flag automatically):

```powershell
.\build.ps1
```

## Usage

```bash
# Run the ASP application in the current directory
axonhta.exe

# Specify the application directory
axonhta.exe --app ./myapp

# Specify window title and size
axonhta.exe --app ./myapp --title "My Tool" --width 1200 --height 800

# Specify a fixed port (random by default)
axonhta.exe --app ./myapp --port 8080

# Map a virtual path to a real directory (repeatable)
axonhta.exe --app ./myapp --alias /music/=D:\Music --alias /photos/=E:\Pictures
```

### Virtual Path Aliases

AxonHTA supports mapping URL path prefixes to real filesystem directories outside the application folder. This is useful for applications that need to access user files (e.g., a music player accessing the user's music library).

Aliases can be configured in two ways:

1. **Command-line flag** (repeatable):

   ```bash
   axonhta.exe --app ./myapp --alias /music/=D:\Music
   ```

2. **Config file** (`data/path_aliases.dat` inside the app directory):

   ```
   ; Virtual path aliases (one per line)
   ; Format: /url-prefix/|C:\real\path
   /music/|C:\Users\jeffr\Music
   /photos/|E:\Pictures
   ```

   The config file is hot-reloaded every 500ms, so applications can write new aliases at runtime (e.g., a settings form) without restarting.

Files served through aliases are protected against path traversal — requests containing `..` are rejected.

### File System Access

HTA applications are trusted desktop applications. The `Scripting.FileSystemObject` has **unrestricted filesystem access** — it can read and write any path on the system, not just the `--app` directory. This allows applications to:

- Scan user-selected directories for media files
- Read and write configuration files anywhere on disk
- Access network shares and external drives

## Application Directory Structure

```
myapp/
├── index.hta       ← entry file (can also be default.hta / index.asp / default.asp)
├── style.css        ← stylesheets
├── data/
│   └── ...          ← runtime data (config, state, etc.)
├── images/
│   └── logo.png
└── include/
    └── common.inc   ← #include files
```

## Development Examples

All sample apps are located in the `HTAtest/` directory and come with a pre-built `axonhta.exe`.

### HTAtest/axonhta-todo-app/

A to-do list application demonstrating **classic ASP page-refresh architecture**.

- Full CRUD: add, edit, toggle complete, delete tasks
- File-based storage (`data/tasks.dat`) using FSO — no database required
- Priority levels (low / medium / high) with color-coded badges
- Filter tabs (all / active / completed) with live counts
- Pure VBScript + HTML + CSS, zero JavaScript framework dependency
- `#include` helper module for reusable CRUD functions

### HTAtest/axonhta-music-player/

A music player built with **HTMX + Alpine.js** on top of AxonHTA.

- HTMX for partial page updates (rescan, settings forms) — zero custom JS for UI
- Alpine.js (~70 lines) for audio playback control only (play/pause/seek/volume/next/prev)
- Virtual path aliases for accessing user's music directory outside the app folder
- FSO to scan directories and serve media files in real time (FSO cache disabled)
- Runtime configuration via `data/path_aliases.dat` (hot-reloaded every 500ms)
- Playback state persistence (`data/state.dat`)
- Dark theme UI with album artwork animation

## Technical Details

- **HTTP server**: Listens on a random port on `127.0.0.1` (not accessible externally).
- **Window**: Launches Chrome/Edge/Chromium in app mode (`--app=<url>`) with a dedicated browser profile per application, isolated from the user's browser session.
- **Browser lookup**: Checks common installation paths on Windows and macOS, and searches `PATH` on Linux. Falls back to the system default browser if none is found.
- **Default page lookup order**: `index.hta` → `default.hta` → `index.asp` → `default.asp` → `index.html` → `default.html`.
- **Session**: Each application instance has an independent in-memory Session.
- **Application**: Application-level global state persists for the lifetime of the window.
- **Server.MapPath**: Maps to the directory specified by `--app`.
- **Script timeout**: Default 90 seconds.
- **HTA tag support**: `windowstate="maximize"` is mapped to `--start-maximized`. Other attributes (`caption`, `icon`, etc.) are logged as unsupported in app mode.
- **FSO cache**: Disabled for HTA apps — all filesystem operations read directly from disk, ensuring real-time updates when files are added or removed.
- **HTML injection**: A small script is automatically injected into every HTML response (including ASP-generated pages) to disable the context menu, prevent drag-and-drop, and send heartbeats for process lifecycle management.

## Comparison with Traditional HTA

| Feature | HTA (traditional) | AxonHTA |
|---------|-------------------|---------|
| Development language | VBScript + HTML + CSS | VBScript + HTML + CSS |
| Rendering engine | IE (deprecated) | Chromium (Chrome/Edge) |
| Cross-platform | Windows only | Windows / macOS / Linux |
| COM calls | Direct `CreateObject` | AxonASP built-in COM compatibility layer |
| Deployment | Single .hta file | Single axonhta.exe + application folder |
| File system access | Full trust | Full trust (trusted desktop application) |
| Virtual path mapping | N/A | `--alias` flag or `data/path_aliases.dat` |
| Partial page updates | N/A (full refresh) | HTMX supported (zero-JS partial swaps) |
| CGO required | N/A | No — pure Go |

## Limitations

- Each interaction without HTMX is a full page refresh (same as classic ASP); partial AJAX updates require HTMX or minimal custom JS.
- The desktop window has no native menu bar / toolbar (must be simulated with HTML).
- `caption="no"` (borderless window) and per-app `icon` from the HTA tag are not supported in app mode.
- Requires a Chromium-based browser to be installed on the system.

## License

Follows the AxonASP main project (MPL-2.0).
