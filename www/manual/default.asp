<%
Dim page, mdPath, mdContent, htmlContent, menuContent, menuHtml, apiAction
Dim g3md, fso, ax, indexPathDir, indexCompiledPath

Set fso = Server.CreateObject("Scripting.FileSystemObject")
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

' Initialize index paths
indexPathDir = Server.MapPath("search-index")
indexCompiledPath = indexPathDir & ax.AxDirSeparator() & "manual"

apiAction = LCase(Trim(Request.QueryString("api")))
If apiAction = "search" Then
    Response.ContentType = "application/json"
    Response.Charset = "utf-8"
    Response.Write SearchManualJson(Trim(Request.QueryString("q")))
    Response.End
End If
'Future implementation
If apiAction = "triggerindexbuild" Then
    Response.ContentType = "application/json"
    Response.Charset = "utf-8"
    InitializeIndexIfNeeded()
    Response.Write GetIndexStatusJson()
    Response.End
End If

If apiAction = "indexstatus" Then
    Response.ContentType = "application/json"
    Response.Charset = "utf-8"
    Response.Write GetIndexStatusJson()
    Response.End
End If

Function BoolToJson(value)
    If CBool(value) Then
        BoolToJson = "true"
    Else
        BoolToJson = "false"
    End If
End Function

Function GetIndexStatusJson()
    Dim indexExists, isBuilding
    indexExists = fso.FolderExists(indexCompiledPath)
    isBuilding = False
    GetIndexStatusJson = "{""exists"":" & BoolToJson(indexExists) & ",""building"":" & BoolToJson(isBuilding) & "}"
End Function


Function InitializeIndexIfNeeded()
    ' Ensure index directory exists and initialize if needed
    Dim search
    On Error Resume Next
    
    If Not fso.FolderExists(indexPathDir) Then
        fso.CreateFolder indexPathDir
    End If
    
    ' Trigger BuildIndex only when compiled index does not exist
    If Not fso.FolderExists(indexCompiledPath) Then
        Set search = Server.CreateObject("G3SEARCH")
        If Err.Number = 0 Then
            search.IndexPath = indexCompiledPath
            search.DocsPath = Server.MapPath("md")
            search.Extension = ".md"
            search.BuildIndex()
            Err.Clear
        End If

        Set search = Nothing
    End If
    
    On Error Goto 0
End Function

' 1. Get requested page

page = Request.QueryString("page")
If page = "" Then page = "md/axonasp/welcome"

' Security: Basic path sanitization
page = Replace(page, "..", "")

' Accept menu links that include .md and normalize to internal page key
If LCase(Right(page, 3)) = ".md" Then
    page = Left(page, Len(page) - 3)
End If

' 2. Load Content
mdPath = Server.MapPath(page & ".md")
If fso.FileExists(mdPath) Then
    mdContent = ReadFile(mdPath)
Else
    mdContent = "# 404 - Page Not Found" & vbCrLf & "The requested documentation page '" & ax.AxStripTags(page) & "' was not found."
End If

' 3. Render Markdown
Set g3md = Server.CreateObject("G3MD")
g3md.Unsafe = True
htmlContent = g3md.Process(mdContent)
If htmlContent = "" And mdContent <> "" Then
    htmlContent = "<p style='color:red'>Error: Markdown rendering failed.</p>"
End If

' 4. Render Menu
Dim menuMdPath
menuMdPath = Server.MapPath("menu.md")
If fso.FileExists(menuMdPath) Then
    menuContent = ReadFile(menuMdPath)
    ' We use a simple custom parser for the menu to generate the tree structure
    menuHtml = ParseMenuToTree(menuContent)
Else
    menuHtml = "Menu not found."
End If

Function ReadFile(path)
    Dim stream
    ReadFile = ""
    On Error Resume Next

    Set stream = Server.CreateObject("ADODB.Stream")
    If Err.Number = 0 Then
        stream.Type = 2 ' Text
        stream.Charset = "utf-8"
        stream.Open

        If Err.Number = 0 Then
            stream.LoadFromFile path
            If Err.Number = 0 Then
                ReadFile = stream.ReadText
            End If
        End If
    End If

    Err.Clear
    If Not stream Is Nothing Then
        stream.Close
        Set stream = Nothing
    End If
    On Error Goto 0
End Function

Function ParseMenuToTree(content)
    ParseMenuToTree = "<ul class='treeview' id='menu-tree'></ul><script type='text/plain' id='menu-md-source'>" & Server.HTMLEncode(content) & "</script>"
End Function


Function SearchManualJson(term)
    Dim search, docsPath, results, resultRow, resultPath, relPath
    Dim i, rowCount, jsonParts

    SearchManualJson = "[]"

    If term = "" Then
        Exit Function
    End If

    On Error Resume Next

    ' Ensure index is initialized (with concurrency protection)
    InitializeIndexIfNeeded()

    Set search = Server.CreateObject("G3SEARCH")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If

    docsPath = Server.MapPath("md")

    search.IndexPath = indexCompiledPath
    search.DocsPath = docsPath
    search.Extension = ".md"

    results = search.Search(term)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If

    If Not IsArray(results) Then
        Exit Function
    End If

    rowCount = UBound(results) - LBound(results) + 1
    If rowCount <= 0 Then
        Exit Function
    End If

    ReDim jsonParts(rowCount - 1)
    rowCount = 0

    For i = LBound(results) To UBound(results)
        resultPath = ""
        resultRow = results(i)

        If IsArray(resultRow) Then
            If UBound(resultRow) >= 0 Then
                resultPath = CStr(resultRow(0))
            End If
        Else
            resultPath = CStr(resultRow)
        End If

        relPath = NormalizeSearchDocPath(resultPath, docsPath)
        If relPath <> "" Then
            jsonParts(rowCount) = Chr(34) & JsonEscape(relPath) & Chr(34)
            rowCount = rowCount + 1
        End If
    Next

    If rowCount <= 0 Then
        Exit Function
    End If

    ReDim Preserve jsonParts(rowCount - 1)
    SearchManualJson = "[" & Join(jsonParts, ",") & "]"
End Function

Function NormalizeSearchDocPath(rawPath, docsPath)
    Dim normalizedPath, normalizedDocs, relPath

    normalizedPath = Replace(CStr(rawPath), "\", "/")
    normalizedDocs = Replace(CStr(docsPath), "\", "/")

    relPath = normalizedPath
    If Len(normalizedDocs) > 0 Then
        If LCase(Left(normalizedPath, Len(normalizedDocs))) = LCase(normalizedDocs) Then
            relPath = Mid(normalizedPath, Len(normalizedDocs) + 1)
            If Left(relPath, 1) = "/" Then
                relPath = Mid(relPath, 2)
            End If
        End If
    End If

    If LCase(Left(relPath, 3)) = "md/" Then
        relPath = Mid(relPath, 4)
    End If

    NormalizeSearchDocPath = relPath
End Function

Function JsonEscape(value)
    Dim escaped
    escaped = CStr(value)
    escaped = Replace(escaped, "\", "\\")
    escaped = Replace(escaped, Chr(34), "\" & Chr(34))
    escaped = Replace(escaped, vbCrLf, "\n")
    escaped = Replace(escaped, vbCr, "\n")
    escaped = Replace(escaped, vbLf, "\n")
    escaped = Replace(escaped, vbTab, "\t")
    JsonEscape = escaped
End Function

%>
<!DOCTYPE html>
<html lang="en">
    <!--
        
        AxonASP Server
        Copyright (C) 2026 G3pix Ltda. All rights reserved.
        
        Developed by Lucas Guimarães - G3pix Ltda
        Contact: https://g3pix.com.br/
        Project URL: https://g3pix.com.br/axonasp
        
        This Source Code Form is subject to the terms of the Mozilla Public
        License, v. 2.0. If a copy of the MPL was not distributed with this
        file, You can obtain one at https://mozilla.org/MPL/2.0/.
        
        Attribution Notice:
        If this software is used in other projects, the name "AxonASP Server"
        must be cited in the documentation or "About" section.
        
        Contribution Policy:
        Modifications to the core source code of AxonASP Server must be
        made available under this same license terms.
        
        -->
    <head>
        <meta charset="UTF-8" />
        <title>
            AxonASP Documentation Library -<%= page %>
        </title>
        <style>
            :root {
                --win-blue-dark: #003399;
                --win-blue-light: #3366cc;
                --win-blue-soft: #c7d7f8;
                --win-bg: #ece9d8;
                --win-border: #808080;
                --win-text: #0f0f0f;
                --win-muted: #404040;
                --win-link: #003399;
                --win-link-hover: #335ea8;
                --win-gold: #ffd700;
                --win-gold-dark: #c8a200;
                --radius-sm: 6px;
                --radius-md: 10px;
                --radius-lg: 14px;
                --shadow-card:
                    0 4px 16px rgba(0, 51, 153, 0.1),
                    0 2px 6px rgba(0, 0, 0, 0.07);
            }

            *,
            *::before,
            *::after {
                box-sizing: border-box;
            }

            html,
            body {
                margin: 0;
                padding: 0;
                height: 100%;
                overflow: hidden;
                font-family: Tahoma, Verdana, Arial, sans-serif;
                font-size: 12px;
                color: var(--win-text);
                background-color: var(--win-bg);
                background-image:
                    radial-gradient(
                        ellipse at 10% 15%,
                        rgba(51, 102, 204, 0.1),
                        transparent 38%
                    ),
                    radial-gradient(
                        ellipse at 88% 8%,
                        rgba(0, 51, 153, 0.07),
                        transparent 32%
                    ),
                    linear-gradient(
                        180deg,
                        #f2efe4 0%,
                        #ece9d8 40%,
                        #e4e0cc 100%
                    );
                background-repeat: no-repeat;
                background-size: 100% 100%;
            }

            /* ── Header ─────────────────────────────────────────── */
            #header {
                background: linear-gradient(
                    90deg,
                    var(--win-blue-dark) 0%,
                    #1f56bc 42%,
                    var(--win-blue-light) 100%
                );
                color: #fff;
                padding: 0 15px;
                height: 60px;
                display: flex;
                align-items: center;
                border-bottom: 3px solid var(--win-blue-light);
                box-shadow: 0 4px 14px rgba(0, 0, 0, 0.2);
                z-index: 100;
            }

            #header h1 {
                font-family: Tahoma, Verdana, serif;
                font-style: normal;
                font-size: 24px;
                margin: 0 0 0 12px;
                font-weight: normal;
                color: #fff;
                text-shadow: 1px 1px 0 rgba(0, 0, 0, 0.35);
            }

            #header .logo {
                margin-right: 3px;
                flex-shrink: 0;
            }

            /* ── Main Layout ─────────────────────────────────────── */
            #main-container {
                display: flex;
                height: calc(100% - 82px);
                border-top: 1px solid #fff;
            }

            /* ── Sidebar ─────────────────────────────────────────── */
            #sidebar {
                width: 300px;
                background: linear-gradient(180deg, #eceae0 0%, #e2e0d6 100%);
                border-right: 1px solid var(--win-border);
                overflow-y: auto;
                padding: 10px;
                font-size: 12px;
                flex-shrink: 0;
            }

            #sidebar .section-title {
                padding: 5px 0;
                margin-top: 15px;
                margin-bottom: 10px;
                font-weight: bold;
                color: #1a3470;
                border-bottom: 2px solid var(--win-blue-light);
                text-transform: uppercase;
                font-size: 11px;
                letter-spacing: 0.4px;
            }

            .sidebar-tabs {
                display: flex;
                border-bottom: 1px solid #aca899;
                margin-bottom: 10px;
            }

            .sidebar-tab-btn {
                flex: 1;
                border: 1px solid #aca899;
                border-bottom: none;
                background: linear-gradient(180deg, #f5f3ea 0%, #e5e2d8 100%);
                color: #1a3470;
                font-family: Tahoma, Verdana, sans-serif;
                font-size: 11px;
                font-weight: bold;
                padding: 6px 8px;
                cursor: pointer;
            }

            .sidebar-tab-btn + .sidebar-tab-btn {
                border-left: none;
            }

            .sidebar-tab-btn.active {
                background: #fff;
                color: var(--win-blue-dark);
            }

            .sidebar-tab-panel {
                display: none;
            }

            .sidebar-tab-panel.active {
                display: block;
            }

            .sidebar-search-input {
                width: 100%;
                box-sizing: border-box;
                font-size: 11px;
                font-family: Tahoma, Verdana, sans-serif;
                border: 1px solid #aca899;
                border-radius: 4px;
                padding: 3px 6px;
                background: #fff;
                transition: border-color 0.15s, box-shadow 0.15s;
            }

            .search-results {
                margin-top: 10px;
                border-top: 1px solid #d2d0c8;
                padding-top: 8px;
            }

            .search-result-item {
                padding: 6px 4px;
                border-bottom: 1px dotted #b9b6ab;
            }

            .search-result-link {
                display: block;
                color: #001f66;
                text-decoration: none;
                font-weight: bold;
                margin-bottom: 2px;
            }

            .search-result-link:hover {
                color: var(--win-blue-dark);
                text-decoration: underline;
            }

            .search-result-folder {
                color: #5f5b4e;
                font-size: 11px;
            }

            .search-empty,
            .search-error {
                color: #5f5b4e;
                font-size: 11px;
                padding: 6px 2px;
            }

            .search-error {
                color: #8b0000;
            }

            #sidebar a {
                color: #111;
                text-decoration: none;
                display: block;
                padding: 2px 4px;
            }

            #sidebar a:hover {
                color: var(--win-blue-dark);
                text-decoration: underline;
            }

            /* ── Content Area ────────────────────────────────────── */
            #content {
                flex: 1;
                background-color: #fff;
                overflow-y: auto;
                padding: 20px 40px;
            }

            /* ── Treeview — preserved exactly, colors updated ────── */
            .treeview,
            .treeview ul {
                list-style-type: none;
                padding-left: 15px;
                margin: 0;
            }

            .treeview li {
                margin: 2px 0;
                white-space: nowrap;
            }

            .treeview li.folder > .folder-toggle {
                cursor: pointer;
                padding-left: 2px;
                display: block;
                position: relative;
            }

            .treeview li.folder > .folder-toggle::before {
                content: "+";
                display: inline-block;
                width: 9px;
                height: 9px;
                border: 1px solid #808080;
                line-height: 8px;
                text-align: center;
                margin-right: 5px;
                background: #fff;
                font-family: "Courier New", monospace;
                font-weight: bold;
                font-size: 10px;
                vertical-align: middle;
            }

            .treeview li.folder.expanded > .folder-toggle::before {
                content: "-";
            }

            .treeview li.folder > ul.submenu {
                display: none;
                padding-left: 15px;
                border-left: 1px dotted #aca899;
                margin-left: 7px;
            }

            .treeview li.folder.expanded > ul.submenu {
                display: block;
            }

            .treeview li.file {
                padding-left: 16px;
                position: relative;
            }

            .treeview li.file::before {
                content: "";
                position: absolute;
                left: -15px;
                top: 10px;
                width: 31px;
                border-top: 1px dotted #aca899;
            }

            .treeview a {
                color: #000;
                text-decoration: none;
                padding: 1px 2px;
                display: inline-block;
                border-radius: 3px;
            }

            .treeview a:hover:not(.selected-node) {
                color: var(--win-blue-dark);
                text-decoration: underline;
            }

            .selected-node,
            .treeview a.selected-node {
                background-color: var(--win-blue-dark) !important;
                color: #fff !important;
                text-decoration: none;
            }

            /* ── Content Typography ──────────────────────────────── */
            #content h1 {
                font-family: Tahoma, Verdana, sans-serif;
                color: var(--win-blue-dark);
                font-size: 22px;
                border-bottom: 3px solid var(--win-blue-light);
                padding-bottom: 6px;
                margin-top: 0;
                margin-bottom: 15px;
            }

            #content h2 {
                font-family: Tahoma, Verdana, sans-serif;
                color: var(--win-blue-dark);
                font-size: 16px;
                margin-top: 25px;
                border-bottom: 1px solid #c0c8d8;
                padding-bottom: 3px;
                margin-bottom: 10px;
            }

            #content h3 {
                font-family: Tahoma, Verdana, sans-serif;
                color: #0e2f78;
                font-size: 14px;
                margin-top: 18px;
                margin-bottom: 7px;
            }

            #content p,
            #content li {
                line-height: 1.6;
                font-size: 12px;
                color: #333;
            }

            #content ul,
            #content ol {
                padding-left: 20px;
                margin-bottom: 12px;
            }

            #content ul {
                list-style: disc;
            }

            #content ol {
                list-style: decimal;
            }

            #content pre {
                background-color: #f0f3f8;
                border-left: 4px solid var(--win-blue-light);
                border-right: 1px solid #ccc;
                border-top: 1px solid #ccc;
                border-bottom: 1px solid #ccc;
                border-radius: 0 var(--radius-sm) var(--radius-sm) 0;
                padding: 12px 14px;
                overflow-x: auto;
                font-family: "Courier New", Courier, monospace;
                font-size: 12px;
                line-height: 1.6;
                margin: 15px 0;
            }

            #content code {
                font-family: "Courier New", Courier, monospace;
                background: rgba(0, 51, 153, 0.07);
                border: 1px solid rgba(0, 51, 153, 0.14);
                border-radius: 3px;
                padding: 1px 5px;
                font-size: 11px;
            }

            #content pre code {
                background: none;
                border: none;
                padding: 0;
                font-size: 12px;
            }

            /* ── Tables ──────────────────────────────────────────── */
            #content table {
                border-collapse: collapse;
                width: 100%;
                margin: 15px 0;
                font-size: 12px;
            }

            #content table th,
            #content table td {
                border: 1px solid #aca899;
                padding: 8px 10px;
                text-align: left;
            }

            #content table th {
                background: linear-gradient(
                    180deg,
                    #1c47a8 0%,
                    var(--win-blue-dark) 100%
                );
                font-weight: bold;
                color: #fff;
            }

            #content table tr:nth-child(even) td {
                background-color: #f4f6fa;
            }

            #content table tr:hover td {
                background-color: #edf3ff;
                color: #0a1f55;
            }

            /* ── Blockquote ──────────────────────────────────────── */
            blockquote {
                margin: 14px 0;
                padding: 8px 14px;
                background: linear-gradient(135deg, #eaf0fc 0%, #dce9ff 100%);
                border: 1px solid #8097c4;
                border-left: 4px solid var(--win-blue-dark);
                border-radius: 0 var(--radius-sm) var(--radius-sm) 0;
                color: #001a4d;
            }

            /* ── Status Bar ──────────────────────────────────────── */
            #status-bar {
                height: 22px;
                background-color: var(--win-bg);
                border-top: 1px solid #aca899;
                font-size: 11px;
                padding: 0 10px;
                display: flex;
                align-items: center;
                color: #000;
            }

            /* ── Index Loading Modal ─────────────────────────────── */
            #index-loading-overlay {
                display: none;
                position: fixed;
                top: 0;
                left: 0;
                right: 0;
                bottom: 0;
                background-color: rgba(0, 0, 0, 0.45);
                z-index: 1000;
                align-items: center;
                justify-content: center;
            }

            #index-loading-overlay.show {
                display: flex;
            }

            .index-loading-modal {
                background-color: var(--win-bg);
                border: 1px solid #8fa8d4;
                border-radius: var(--radius-md);
                box-shadow: 0 6px 24px rgba(0, 51, 153, 0.18), 0 2px 8px rgba(0, 0, 0, 0.10);
                min-width: 400px;
                overflow: hidden;
            }

            .index-loading-header {
                background: linear-gradient(90deg, var(--win-blue-dark) 0%, var(--win-blue-light) 100%);
                color: #fff;
                padding: 8px 15px;
                font-weight: bold;
                font-size: 12px;
            }

            .index-loading-body {
                padding: 25px;
                text-align: center;
                color: #0f0f0f;
                font-size: 12px;
            }

            .index-loading-spinner {
                display: inline-block;
                width: 32px;
                height: 32px;
                border: 3px solid #c0c8d8;
                border-top-color: var(--win-blue-dark);
                border-radius: 50%;
                animation: spin 0.8s linear infinite;
                margin-bottom: 15px;
            }

            .index-loading-message {
                font-size: 13px;
                color: #0f0f0f;
                line-height: 1.5;
            }

            @keyframes spin {
                to {
                    transform: rotate(360deg);
                }
            }
        </style>
    </head>
    <body>
        <div id="header">
            <div class="logo">
                <img
                    src="<%= ax.AxGetLogo() %>"
                    alt="AxonASP"
                    width="43"
                />
            </div>
            <h1>AxonASP Server Documentation Library</h1>
        </div>

        <div id="main-container">
            <div id="sidebar">
                <div class="sidebar-tabs">
                    <button
                        type="button"
                        class="sidebar-tab-btn active"
                        data-tab-target="contents"
                    >
                        Contents
                    </button>
                    <button
                        type="button"
                        class="sidebar-tab-btn"
                        data-tab-target="search"
                    >
                        Search
                    </button>
                </div>

                <div
                    id="sidebar-tab-contents"
                    class="sidebar-tab-panel active"
                    data-tab-panel="contents"
                >
                    <div style="margin-bottom: 10px">
                        <input
                            type="text"
                            id="search-input"
                            placeholder="Search..."
                            class="sidebar-search-input"
                            onfocus="this.style.borderColor='#3366cc';this.style.boxShadow='0 0 0 2px rgba(51,102,204,0.18)'"
                            onblur="
                                this.style.borderColor = '#aca899';
                                this.style.boxShadow = '';
                            "
                        />
                    </div>
                    <%= menuHtml %>
                </div>

                <div
                    id="sidebar-tab-search"
                    class="sidebar-tab-panel"
                    data-tab-panel="search"
                >
                    <div style="margin-bottom: 10px">
                        <input
                            type="text"
                            id="fulltext-search-input"
                            placeholder="Search documentation..."
                            class="sidebar-search-input"
                            onfocus="this.style.borderColor='#3366cc';this.style.boxShadow='0 0 0 2px rgba(51,102,204,0.18)'"
                            onblur="
                                this.style.borderColor = '#aca899';
                                this.style.boxShadow = '';
                            "
                        />
                    </div>
                    <div id="fulltext-search-results" class="search-results">
                        <div class="search-empty">Type to search the documentation index.</div>
                    </div>
                </div>
            </div>
            <div id="content">
                <%= htmlContent %>
            </div>
        </div>

        <div id="status-bar">
            Page:
            <%= ax.AxStripTags(page) %>.md
        </div>

        <!-- Index Loading Modal -->
        <div id="index-loading-overlay">
            <div class="index-loading-modal">
                <div class="index-loading-header">AxonASP Documentation Library</div>
                <div class="index-loading-body">
                    <div class="index-loading-spinner"></div>
                    <div class="index-loading-message">
                        The search index is currently being created.<br />
                        Please wait. This window will close automatically.
                    </div>
                </div>
            </div>
        </div>

        <script>
            // Index Loading Modal Logic
            (function() {
                const overlay = document.getElementById('index-loading-overlay');
                let pollIntervalId = null;

                function checkIndexStatus() {
                    fetch('?api=indexstatus', {
                        method: 'GET',
                        headers: { Accept: 'application/json' }
                    })
                    .then(response => response.json())
                    .then(data => {
                        if (data.exists === true && !data.building) {
                            // Index is ready, close modal
                            if (pollIntervalId) {
                                clearInterval(pollIntervalId);
                                pollIntervalId = null;
                            }
                            overlay.classList.remove('show');
                        }
                    })
                    .catch(error => {
                        console.error('Index status check failed:', error);
                    });
                }

                function initializeIndexModal() {
                    // Check initial status
                    fetch('?api=indexstatus', {
                        method: 'GET',
                        headers: { Accept: 'application/json' }
                    })
                    .then(response => response.json())
                    .then(data => {
                        if (data.exists === true && data.building !== true) {
                            return;
                        }

                        // Show modal immediately while build is running or about to start.
                        overlay.classList.add('show');
                        if (!pollIntervalId) {
                            pollIntervalId = setInterval(checkIndexStatus, 2000);
                        }

                        if (data.building === true) {
                            return;
                        }

                        fetch('?api=triggerindexbuild', {
                            method: 'GET',
                            headers: { Accept: 'application/json' }
                        }).catch(error => {
                            console.error('Index build trigger failed:', error);
                        });
                    })
                    .catch(error => {
                        console.error('Initial index status check failed:', error);
                    });
                }

                // Run when DOM is ready
                if (document.readyState === 'loading') {
                    document.addEventListener('DOMContentLoaded', initializeIndexModal);
                } else {
                    initializeIndexModal();
                }
            })();

            // Tree View Logic
            document.addEventListener("DOMContentLoaded", function () {
                const sidebarTabButtons = document.querySelectorAll(
                    ".sidebar-tab-btn"
                );
                const sidebarTabPanels = document.querySelectorAll(
                    ".sidebar-tab-panel"
                );

                sidebarTabButtons.forEach((button) => {
                    button.addEventListener("click", function () {
                        const target = this.dataset.tabTarget || "contents";

                        sidebarTabButtons.forEach((btn) => {
                            btn.classList.toggle("active", btn === this);
                        });

                        sidebarTabPanels.forEach((panel) => {
                            panel.classList.toggle(
                                "active",
                                panel.dataset.tabPanel === target
                            );
                        });
                    });
                });

                function getIndent(line) {
                    const leading = (line.match(/^[ \t]*/) || [""])[0];
                    return leading.replace(/\t/g, "    ").length;
                }

                function parseMenuMarkdown(markdown) {
                    const lines = markdown
                        .replace(/\r\n/g, "\n")
                        .replace(/\r/g, "\n")
                        .split("\n");

                    const root = {
                        type: "folder",
                        name: "root",
                        indent: -1,
                        children: [],
                    };
                    const stack = [root];

                    lines.forEach((line) => {
                        const trimmed = line.trim();
                        if (!trimmed || trimmed.startsWith("#")) {
                            return;
                        }
                        if (
                            !(
                                trimmed.startsWith("* ") ||
                                trimmed.startsWith("- ")
                            )
                        ) {
                            return;
                        }

                        const indent = getIndent(line);
                        const content = trimmed.slice(2).trim();
                        const match = content.match(
                            /^\[([^\]]+)\]\(([^)]+)\)$/
                        );

                        const node = match
                            ? {
                                  type: "file",
                                  name: match[1],
                                  page: match[2],
                                  indent,
                              }
                            : {
                                  type: "folder",
                                  name: content,
                                  indent,
                                  children: [],
                              };

                        while (
                            stack.length > 1 &&
                            indent <= stack[stack.length - 1].indent
                        ) {
                            stack.pop();
                        }

                        const parent = stack[stack.length - 1];
                        if (!parent.children) {
                            parent.children = [];
                        }
                        parent.children.push(node);

                        if (node.type === "folder") {
                            stack.push(node);
                        }
                    });

                    return root.children || [];
                }

                function renderTree(nodes, container) {
                    nodes.forEach((node) => {
                        const li = document.createElement("li");
                        li.dataset.label = (node.name || "").toLowerCase();

                        if (node.type === "folder") {
                            li.className = "folder collapsed";

                            const toggle = document.createElement("span");
                            toggle.className = "folder-toggle";
                            toggle.textContent = node.name;
                            li.appendChild(toggle);

                            const ul = document.createElement("ul");
                            ul.className = "submenu";
                            li.appendChild(ul);

                            renderTree(node.children || [], ul);
                        } else {
                            li.className = "file";
                            const link = document.createElement("a");
                            link.textContent = node.name;
                            link.href =
                                "?page=" + encodeURIComponent(node.page || "");
                            link.dataset.page = (node.page || "").toLowerCase();
                            li.appendChild(link);
                        }

                        container.appendChild(li);
                    });
                }

                const menuSourceEl = document.getElementById("menu-md-source");
                const menuTreeEl = document.getElementById("menu-tree");
                const menuMarkdown = menuSourceEl
                    ? menuSourceEl.textContent || ""
                    : "";
                const menuNodes = parseMenuMarkdown(menuMarkdown);
                renderTree(menuNodes, menuTreeEl);

                document
                    .getElementById("sidebar")
                    .addEventListener("click", function (event) {
                        const toggle = event.target.closest(".folder-toggle");
                        if (!toggle) {
                            return;
                        }
                        const folder = toggle.parentElement;
                        folder.classList.toggle("expanded");
                        folder.classList.toggle("collapsed");
                    });

                // Search Logic
                const searchInput = document.getElementById("search-input");
                searchInput.addEventListener("input", function () {
                    const filter = this.value.toLowerCase();

                    function filterNode(li) {
                        const ownLabel = (li.dataset.label || "").toLowerCase();
                        const childList = li.querySelector(
                            ":scope > ul.submenu"
                        );

                        if (!childList) {
                            const visible =
                                filter === "" || ownLabel.indexOf(filter) > -1;
                            li.style.display = visible ? "" : "none";
                            return visible;
                        }

                        let hasVisibleChild = false;
                        childList
                            .querySelectorAll(":scope > li")
                            .forEach((childLi) => {
                                if (filterNode(childLi)) {
                                    hasVisibleChild = true;
                                }
                            });

                        const ownMatch =
                            filter === "" || ownLabel.indexOf(filter) > -1;
                        const visible = ownMatch || hasVisibleChild;
                        li.style.display = visible ? "" : "none";

                        if (filter !== "" && hasVisibleChild) {
                            li.classList.add("expanded");
                            li.classList.remove("collapsed");
                        }

                        return visible;
                    }

                    menuTreeEl
                        .querySelectorAll(":scope > li")
                        .forEach((topLi) => {
                            filterNode(topLi);
                        });
                });

                // Expand the current category based on URL
                const currentPage = (
                    new URLSearchParams(window.location.search).get("page") ||
                    "axonasp/welcome"
                ).toLowerCase();
                const links = document.querySelectorAll(".treeview a");
                links.forEach((link) => {
                    const linkPage = (link.dataset.page || "").toLowerCase();
                    if (linkPage === currentPage) {
                        link.classList.add("selected-node");
                        let parent = link.closest(".folder");
                        while (parent) {
                            parent.classList.add("expanded");
                            parent.classList.remove("collapsed");
                            parent = parent.parentElement.closest(".folder");
                        }
                        setTimeout(() => {
                            link.scrollIntoView({
                                behavior: "auto",
                                block: "center",
                            });
                        }, 50);
                    }
                });

                const fulltextSearchInput = document.getElementById(
                    "fulltext-search-input"
                );
                const fulltextSearchResults = document.getElementById(
                    "fulltext-search-results"
                );

                function toTitleCaseFromSlug(value) {
                    const base = String(value || "")
                        .replace(/\.md$/i, "")
                        .replace(/[-_]+/g, " ")
                        .trim();
                    return base.replace(/\b\w/g, (m) => m.toUpperCase());
                }

                function renderSearchResults(paths) {
                    if (!Array.isArray(paths) || paths.length === 0) {
                        fulltextSearchResults.innerHTML =
                            '<div class="search-empty">No results found.</div>';
                        return;
                    }

                    const fragment = document.createDocumentFragment();
                    fulltextSearchResults.innerHTML = "";

                    paths.forEach((rawPath) => {
                        const normalizedPath = String(rawPath || "")
                            .replace(/\\/g, "/")
                            .replace(/^\/+/, "")
                            .replace(/^md\//i, "");

                        if (!normalizedPath) {
                            return;
                        }

                        const parts = normalizedPath.split("/").filter(Boolean);
                        if (parts.length === 0) {
                            return;
                        }

                        const fileName = parts[parts.length - 1];
                        const folderName =
                            parts.length > 1 ? parts[parts.length - 2] : "Manual";
                        const pageTarget = "md/" + normalizedPath;

                        const item = document.createElement("div");
                        item.className = "search-result-item";

                        const link = document.createElement("a");
                        link.className = "search-result-link";
                        link.href =
                            "?page=" + encodeURIComponent(pageTarget);
                        link.textContent = toTitleCaseFromSlug(fileName);

                        const folder = document.createElement("div");
                        folder.className = "search-result-folder";
                        folder.textContent = toTitleCaseFromSlug(folderName);

                        item.appendChild(link);
                        item.appendChild(folder);
                        fragment.appendChild(item);
                    });

                    if (!fragment.childNodes.length) {
                        fulltextSearchResults.innerHTML =
                            '<div class="search-empty">No results found.</div>';
                        return;
                    }

                    fulltextSearchResults.appendChild(fragment);
                }

                let searchDebounceId = 0;
                let activeSearchToken = 0;

                fulltextSearchInput.addEventListener("input", function () {
                    const term = this.value.trim();
                    window.clearTimeout(searchDebounceId);

                    if (term === "") {
                        fulltextSearchResults.innerHTML =
                            '<div class="search-empty">Type to search the documentation index.</div>';
                        return;
                    }

                    searchDebounceId = window.setTimeout(async () => {
                        const token = ++activeSearchToken;
                        fulltextSearchResults.innerHTML =
                            '<div class="search-empty">Searching...</div>';

                        try {
                            const response = await fetch(
                                "?api=search&q=" + encodeURIComponent(term),
                                {
                                    method: "GET",
                                    headers: {
                                        Accept: "application/json",
                                    },
                                }
                            );

                            if (!response.ok) {
                                throw new Error("HTTP " + response.status);
                            }

                            const payload = await response.json();
                            if (token !== activeSearchToken) {
                                return;
                            }

                            renderSearchResults(payload);
                        } catch (error) {
                            if (token !== activeSearchToken) {
                                return;
                            }
                            fulltextSearchResults.innerHTML =
                                '<div class="search-error">Search is temporarily unavailable.</div>';
                        }
                    }, 220);
                });
            });
        </script>
    </body>
</html>
