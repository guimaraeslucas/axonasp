# AxonHTA Runtime Overview

## Overview
AxonHTA is a modern desktop application runtime built for AxonASP that allows developers to create cross-platform desktop applications using pure Classic ASP (VBScript or Server-Side JavaScript), HTML, and CSS. It replaces traditional Internet Explorer-based HTML Applications (HTA) by running an embedded AxonASP web server coupled with system Chromium browsers launched in standalone application mode.

With AxonHTA, developers can build complete desktop user interfaces without requiring complex CGO bindings, heavy Electron runtimes, or client-side JavaScript frameworks.

## Architecture and Execution Model

AxonHTA operates by combining an embedded web server with a local browser instance operating in app mode:

```text
VBScript / JScript (.hta / .asp) -> AxonASP VM -> Internal HTTP Server (127.0.0.1) -> Chromium App Window
```

1. **Embedded Web Server Initialization:** Executing the AxonHTA runtime (`axonhta.exe`) initializes an in-memory AxonASP HTTP server bound to a random local loopback port (`127.0.0.1`).
2. **Script Execution:** Requested `.hta` and `.asp` files are compiled and executed by the internal AxonASP Virtual Machine using standard ASP intrinsics and native libraries.
3. **App-Mode Browser Window:** AxonHTA searches the host operating system for an installed Chromium-based browser (Google Chrome, Microsoft Edge, Chromium, or Brave) and launches it with the `--app` flag, rendering a clean, borderless application window with no browser address bar or tabs.
4. **Lifecycle and Heartbeat:** AxonHTA automatically injects a lightweight heartbeat monitoring script into rendered HTML responses. The browser sends periodic `HEAD` requests (`/__heartbeat__`) to the local server every 5 seconds. When the window is closed, heartbeats cease, and AxonHTA cleanly terminates the server process after a 15-second grace window.
5. **Context Restrictions:** The injected runtime script automatically disables default browser context menus (right-click) and element drag-and-drop actions to maintain a native desktop look and feel.

## Technical Details

- **Local Host Isolation:** The internal HTTP server binds strictly to local loopback (`127.0.0.1`), ensuring the application endpoints are never exposed to external networks.
- **Default Entry Point Resolution:** When pointing AxonHTA to an application folder, the entry page is resolved using the following priority order:
  - `index.hta`
  - `default.hta`
  - `index.asp`
  - `default.asp`
  - `index.html`
  - `default.html`
- **Application State and Session Scope:** Each launched instance of an AxonHTA application maintains independent `Session` and `Application` object memory pools for state persistence across page interactions.
- **Unrestricted FileSystemObject Access:** Because AxonHTA applications run as trusted local desktop processes, the `Scripting.FileSystemObject` library operates without web server sandbox constraints, allowing full read and write access across local drives and network shares.
- **FileSystem Cache Bypass:** Disk caching for `Scripting.FileSystemObject` operations is disabled in AxonHTA mode to ensure real-time file updates are immediately reflected in the user interface.

## Comparison: Traditional HTA vs. AxonHTA

| Feature | Traditional Windows HTA | G3Pix AxonHTA |
| --- | --- | --- |
| Execution Engine | Internet Explorer (Trident / Deprecated) | Modern Chromium (Chrome, Edge, Brave) |
| Operating System Support | Windows Only | Cross-Platform (Windows, Linux, macOS) |
| Scripting Languages | VBScript / JScript | VBScript / JScript via AxonASP VM |
| Packaging & Distribution | Direct `.hta` file | Portable `axonhta.exe` + Application Folder |
| File System Access | Full Trust COM | Full Trust `Scripting.FileSystemObject` |
| Virtual Path Aliasing | Not Supported | Command-Line `--alias` & `data/path_aliases.dat` |
| Partial Page Updates | Full Page Refresh Only | Full Page Refresh or HTMX Partial Swaps |
| CGO Dependency | Not Applicable | Pure Go (No CGO Required) |

## Code Example

Below is a complete, minimal `index.hta` application displaying system details using VBScript and AxonASP native functions:

```asp
<%@ Language="VBScript" %>
<%
Dim axon
Set axon = Server.CreateObject("G3AXON.FUNCTIONS")

Dim sysInfo, osName, userHome
sysInfo = axon.Axsysteminfo()
osName = axon.Axgetenv("OS")
userHome = axon.Axuserhomedirpath()
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>AxonHTA System Info</title>
    <link rel="stylesheet" href="/css/axonasp.css">
</head>
<body style="padding: 20px;">
    <div class="card">
        <h3>AxonHTA Application Environment</h3>
        <p><strong>System Info:</strong> <%= sysInfo %></p>
        <p><strong>Operating System:</strong> <%= osName %></p>
        <p><strong>User Home Directory:</strong> <%= userHome %></p>
    </div>
</body>
</html>
```
