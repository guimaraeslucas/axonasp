# Building and Running AxonHTA Applications

## Overview
This guide documents the procedures for compiling the AxonHTA executable, executing desktop applications, configuring command-line flags, and mapping virtual path aliases.

## Prerequisites
Before compiling or running AxonHTA, ensure your environment meets the following requirements:
- Go compiler version 1.22 or higher installed.
- AxonASP repository cloned locally.
- A Chromium-based web browser installed on the target system (Google Chrome, Microsoft Edge, Chromium, or Brave).

## Building the AxonHTA Executable

AxonHTA is written in pure Go without CGO dependencies. To compile the executable from source code:

1. Open a terminal and navigate to the `axonhta` source directory:
```bash
cd axonhta
```

2. Compile the binary using the Go toolchain:
```bash
go build -o axonhta.exe
```

### Hiding the Windows Console Window
By default, launching `axonhta.exe` on Windows displays a command prompt window behind the application GUI. To build a GUI-only binary that runs silently in the background, append the `-H windowsgui` linker flag:

```bash
go build -ldflags="-H windowsgui" -o axonhta.exe
```

Alternatively, run the included PowerShell build script located in `./axonhta/build.ps1`, which automatically applies the GUI linker flag:

```powershell
.\build.ps1
```

## Running Applications and Command-Line Flags

AxonHTA accepts command-line parameters to control application target folders, window dimensions, titles, network ports, and virtual path mappings:

```bash
# Launch the application in the current directory
axonhta.exe

# Specify a target application folder
axonhta.exe --app ./myapp

# Custom window title and startup dimensions
axonhta.exe --app ./myapp --title "Customer Dashboard" --width 1280 --height 768

# Bind to a specific local HTTP port (random port assigned by default)
axonhta.exe --app ./myapp --port 8080

# Map virtual path aliases for external directories
axonhta.exe --app ./myapp --alias /music/=D:\Media\Music --alias /photos/=E:\Pictures
```

## Virtual Path Aliases

AxonHTA supports mapping virtual URL prefixes to real filesystem directories outside the main `--app` directory. This allows desktop applications to browse, read, and process user files across local or network drives.

Path aliases can be defined using two methods:

### 1. Command-Line Flag
Pass one or more `--alias` flags when executing the runtime:
```bash
axonhta.exe --app ./myapp --alias /documents/=C:\Users\Public\Documents
```

### 2. Hot-Reloaded Configuration File
Create or update the text file located at `data/path_aliases.dat` inside the application directory:
```text
; Virtual path aliases (Format: /url-prefix/|C:\real\path)
/music/|D:\UserMedia\Music
/photos/|E:\Archive\Photos
```
AxonHTA monitors `data/path_aliases.dat` and automatically reloads alias definitions every 500 milliseconds. Applications can dynamically write new aliases at runtime without restarting the process.

All file requests served through virtual path aliases include strict security validation to prevent path traversal attacks; any request containing `..` relative path segments is automatically rejected.

## Standard Application Directory Structure

Organize your AxonHTA project files according to the following recommended layout:

```text
myapp/
├── index.hta        <- Main application entry page
├── style.css        <- Custom application CSS
├── data/            <- Application data storage (hot-reloaded config files, databases)
│   └── path_aliases.dat
├── images/          <- Graphic assets
│   └── logo.png
└── include/         <- Reusable VBScript/JScript include files
    └── common.inc
```

## Code Example

The following example demonstrates building a desktop file scanner page using VBScript and `Scripting.FileSystemObject`:

```asp
<%@ Language="VBScript" %>
<%
Dim fso, targetFolder, folderObj, fileObj, fileCount
Set fso = Server.CreateObject("Scripting.FileSystemObject")

targetFolder = Server.MapPath(".")
fileCount = 0

If fso.FolderExists(targetFolder) Then
    Set folderObj = fso.GetFolder(targetFolder)
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>AxonHTA File Scanner</title>
    <link rel="stylesheet" href="/css/axonasp.css">
</head>
<body style="padding: 20px;">
    <div class="card">
        <h3>Application Directory Contents</h3>
        <p><strong>Path:</strong> <%= targetFolder %></p>
        <table class="table-wrap">
            <thead>
                <tr>
                    <th>File Name</th>
                    <th>Size (Bytes)</th>
                    <th>Last Modified</th>
                </tr>
            </thead>
            <tbody>
            <%
            For Each fileObj In folderObj.Files
                fileCount = fileCount + 1
            %>
                <tr>
                    <td><%= fileObj.Name %></td>
                    <td><%= fileObj.Size %></td>
                    <td><%= fileObj.DateLastModified %></td>
                </tr>
            <%
            Next
            %>
            </tbody>
        </table>
        <p style="margin-top: 10px;"><strong>Total Files:</strong> <%= fileCount %></p>
    </div>
</body>
</html>
<%
End If
Set fso = Nothing
%>
```
