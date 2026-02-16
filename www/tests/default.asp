<%@ Language=VBScript %>
<%
Option Explicit

Dim fso, testsDir, folderObj, fileObj
Dim testFiles(), testCount
Dim currentFile

Set fso = Server.CreateObject("Scripting.FileSystemObject")
testsDir = Server.MapPath("/tests")

testCount = 0
ReDim testFiles(0)

If fso.FolderExists(testsDir) Then
    Set folderObj = fso.GetFolder(testsDir)
    For Each fileObj In folderObj.Files
        If LCase(Right(fileObj.Name, 4)) = ".asp" Then
            If LCase(Left(fileObj.Name, 5)) = "test_" Then
                If LCase(fileObj.Name) <> "default.asp" Then
                    If testCount = 0 Then
                        ReDim testFiles(0)
                    Else
                        ReDim Preserve testFiles(testCount)
                    End If
                    testFiles(testCount) = fileObj.Name
                    testCount = testCount + 1
                End If
            End If
        End If
    Next
End If

If testCount > 1 Then
    Dim i, j, tmp
    For i = 0 To testCount - 2
        For j = i + 1 To testCount - 1
            If LCase(testFiles(i)) > LCase(testFiles(j)) Then
                tmp = testFiles(i)
                testFiles(i) = testFiles(j)
                testFiles(j) = tmp
            End If
        Next
    Next
End If

currentFile = Request.QueryString("file")
If currentFile = "" And testCount > 0 Then
    currentFile = testFiles(0)
End If
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AxonASP Tests Dashboard</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        html, body { height: 100%; }
        body {
            font-family: Segoe UI, Tahoma, Arial, sans-serif;
            background: #f3f4f6;
        }
        .layout {
            display: flex;
            height: 100vh;
            width: 100%;
        }
        .sidebar {
            width: 340px;
            background: #111827;
            color: #e5e7eb;
            border-right: 1px solid #1f2937;
            display: flex;
            flex-direction: column;
        }
        .sidebar-header {
            padding: 16px;
            border-bottom: 1px solid #1f2937;
        }
        .sidebar-header h1 {
            font-size: 18px;
            margin-bottom: 4px;
        }
        .sidebar-header p {
            color: #9ca3af;
            font-size: 13px;
        }
        .tests-list {
            padding: 12px;
            overflow-y: auto;
            flex: 1;
        }
        .test-btn {
            display: block;
            width: 100%;
            text-align: left;
            padding: 10px 12px;
            margin-bottom: 8px;
            border: 1px solid #374151;
            border-radius: 8px;
            background: #1f2937;
            color: #e5e7eb;
            text-decoration: none;
            font-size: 13px;
            transition: background .15s ease, border-color .15s ease;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .test-btn:hover {
            background: #273549;
            border-color: #4b5563;
        }
        .test-btn.active {
            background: #2563eb;
            border-color: #1d4ed8;
            color: #fff;
        }
        .viewer {
            flex: 1;
            display: flex;
            flex-direction: column;
            min-width: 0;
            background: #fff;
        }
        .viewer-header {
            height: 54px;
            border-bottom: 1px solid #e5e7eb;
            display: flex;
            align-items: center;
            padding: 0 16px;
            font-weight: 600;
            color: #111827;
            background: #f9fafb;
        }
        .viewer-body {
            flex: 1;
            min-height: 0;
        }
        iframe {
            width: 100%;
            height: 100%;
            border: 0;
            background: #fff;
        }
        .empty {
            height: 100%;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #6b7280;
            font-size: 15px;
        }
    </style>
</head>
<body>
    <div class="layout">
        <aside class="sidebar">
            <div class="sidebar-header">
                <h1>AxonASP Tests</h1>
                <p><%= testCount %> available test files</p>
            </div>
            <div class="tests-list">
                <% If testCount = 0 Then %>
                    <div>No test_*.asp files found in /tests.</div>
                <% Else
                    Dim index
                    For index = 0 To testCount - 1
                        Dim fileName, activeClass
                        fileName = testFiles(index)
                        activeClass = ""
                        If LCase(fileName) = LCase(currentFile) Then
                            activeClass = " active"
                        End If
                %>
                    <a
                        class="test-btn<%= activeClass %>"
                        href="<%= fileName %>"
                        data-file="<%= fileName %>"
                        onclick="return loadTest(this);"
                    ><%= fileName %></a>
                <%  Next
                   End If %>
            </div>
        </aside>

        <section class="viewer">
            <div class="viewer-header" id="current-file-name"><%= currentFile %></div>
            <div class="viewer-body">
                <% If currentFile = "" Then %>
                    <div class="empty">Select a test from the left menu.</div>
                <% Else %>
                    <iframe id="test-frame" src="<%= currentFile %>" title="<%= currentFile %>"></iframe>
                <% End If %>
            </div>
        </section>
    </div>
    <script>
        function loadTest(element) {
            if (!element) {
                return false;
            }

            var fileName = element.getAttribute("data-file");
            if (!fileName) {
                return false;
            }

            var frame = document.getElementById("test-frame");
            var fileTitle = document.getElementById("current-file-name");

            if (frame) {
                frame.src = fileName;
                frame.title = fileName;
            }

            if (fileTitle) {
                fileTitle.textContent = fileName;
            }

            var buttons = document.querySelectorAll(".test-btn");
            for (var index = 0; index < buttons.length; index++) {
                buttons[index].classList.remove("active");
            }
            element.classList.add("active");

            if (window.history && window.history.replaceState) {
                var nextUrl = "default.asp?file=" + encodeURIComponent(fileName);
                window.history.replaceState({ file: fileName }, "", nextUrl);
            }

            return false;
        }
    </script>
</body>
</html>