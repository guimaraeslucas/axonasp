# AxonHTA Sample Applications and Development Patterns

## Overview
AxonHTA includes complete reference applications located in the repository at `./axonhta/HTAtest`. These sample applications demonstrate key development patterns including classic page-refresh CRUD architectures, HTMX partial rendering, file-based persistence, and media streaming via path aliases.

## Repository Sample Applications

Developers can inspect, run, and modify the pre-built demonstration applications located directly inside the project repository under `./axonhta/HTAtest/`:

- `./axonhta/HTAtest/axonhta-todo-app/`: A classic task management application built with pure VBScript and FileSystemObject persistence.
- `./axonhta/HTAtest/axonhta-music-player/`: A modern desktop audio player utilizing HTMX, Alpine.js, and virtual path aliases for local media libraries.

Each sample directory includes a pre-compiled `axonhta.exe` executable for immediate testing.

## Sample 1: Task Manager (`axonhta-todo-app`)

The To-Do application located in `./axonhta/HTAtest/axonhta-todo-app/` showcases a complete desktop CRUD interface built without external JavaScript dependencies:

- **Storage:** Data is persisted in plain text format (`data/tasks.dat`) using `Scripting.FileSystemObject`.
- **Architecture:** Classic page-refresh pattern where HTML forms submit standard POST requests to ASP handlers.
- **Features:** Create, read, update status, delete tasks, filter by state (all, active, completed), and display priority badges.
- **Code Organization:** Modular design utilizing `#include` directives (`include/common.inc`) to separate file I/O operations from template rendering.

### Task Storage Handler Example (`include/common.inc`)

```asp
<%
Function ReadTasks(filePath)
    Dim fso, file, lines, lineText
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    lines = Array()
    
    If fso.FileExists(filePath) Then
        Set file = fso.OpenTextFile(filePath, 1, False)
        Do Until file.AtEndOfStream
            lineText = file.ReadLine()
            If Len(Trim(lineText)) > 0 Then
                ReDim Preserve lines(UBound(lines) + 1)
                lines(UBound(lines)) = lineText
            End If
        Loop
        file.Close()
    End If
    
    ReadTasks = lines
    Set fso = Nothing
End Function
%>
```

## Sample 2: Reactive Desktop Music Player (`axonhta-music-player`)

The Music Player application located in `./axonhta/HTAtest/axonhta-music-player/` illustrates building highly reactive desktop applications by pairing AxonHTA with HTMX and Alpine.js:

- **Reactive UI:** Uses HTMX (`hx-post`, `hx-target`) for seamless partial DOM updates without full browser reloads.
- **Media Access:** Accesses local audio files outside the application directory by registering virtual path aliases in `data/path_aliases.dat`.
- **Audio Control:** Employs minimal Alpine.js script blocks (~70 lines) strictly for audio element playback controls (play, pause, volume, track seeking).
- **Directory Scanning:** Scans designated music folders via `Scripting.FileSystemObject` and streams media files in real time.
- **State Persistence:** Saves user playback state and settings to `data/state.dat`.

### HTMX Partial Search Rendering Pattern Example

```asp
<%@ Language="VBScript" %>
<%
Dim fso, musicDir, searchQuery, isPartial
searchQuery = Request.Form("query")
isPartial = (Request.ServerVariables("HTTP_HX_REQUEST") = "true")

Set fso = Server.CreateObject("Scripting.FileSystemObject")
musicDir = Server.MapPath("/music/")

' If request originates from HTMX, render only the updated track list table
If isPartial Then
%>
    <div id="trackList">
        <p>Search Results for: <strong><%= Server.HTMLEncode(searchQuery) %></strong></p>
        <!-- Dynamic track rows rendered here -->
    </div>
<%
    Response.End()
End If
Set fso = Nothing
%>
```

## Running the Sample Projects

To execute any of the sample applications:

1. Open a terminal and navigate to the target sample directory:
```bash
cd axonhta/HTAtest/axonhta-todo-app
```

2. Run the bundled executable:
```bash
axonhta.exe
```

3. Alternatively, launch the sample from the repository root specifying the `--app` directory:
```bash
axonhta.exe --app ./axonhta/HTAtest/axonhta-music-player
```
