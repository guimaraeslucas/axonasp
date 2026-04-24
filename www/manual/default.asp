<%
Dim page, mdPath, mdContent, htmlContent, menuContent, menuHtml
Dim g3md, fso

Set fso = Server.CreateObject("Scripting.FileSystemObject")
Set g3md = Server.CreateObject("G3Md")

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
    mdContent = "# 404 - Page Not Found" & vbCrLf & "The requested documentation page '" & page & "' was not found."
End If

' 3. Render Markdown
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
    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 2 ' Text
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    ReadFile = stream.ReadText
    stream.Close
End Function

Function ParseMenuToTree(content)
    ParseMenuToTree = "<ul class='treeview' id='menu-tree'></ul><script type='text/plain' id='menu-md-source'>" & Server.HTMLEncode(content) & "</script>"
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
        </style>
    </head>
    <body>
        <div id="header">
            <div class="logo">
                <%
                Dim ax
                Set ax = Server.CreateObject("G3Axon.Functions")
                %>
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
                <div
                    style="
                        margin-bottom: 10px;
                        border-bottom: 1px solid #aca899;
                        padding-bottom: 5px;
                    "
                >
                    <strong>Contents</strong>
                </div>
                <div style="margin-bottom: 10px">
                    <input
                        type="text"
                        id="search-input"
                        placeholder="Search..."
                        style="
                            width: 100%;
                            box-sizing: border-box;
                            font-size: 11px;
                            font-family: Tahoma, Verdana, sans-serif;
                            border: 1px solid #aca899;
                            border-radius: 4px;
                            padding: 3px 6px;
                            background: #fff;
                            transition:
                                border-color 0.15s,
                                box-shadow 0.15s;
                        "
                        onfocus="this.style.borderColor='#3366cc';this.style.boxShadow='0 0 0 2px rgba(51,102,204,0.18)'"
                        onblur="
                            this.style.borderColor = '#aca899';
                            this.style.boxShadow = '';
                        "
                    />
                </div>
                <%= menuHtml %>
            </div>
            <div id="content">
                <%= htmlContent %>
            </div>
        </div>

        <div id="status-bar">
            Page:
            <%= page %>.md
        </div>

        <script>
            // Tree View Logic
            document.addEventListener("DOMContentLoaded", function () {
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
            });
        </script>
    </body>
</html>
