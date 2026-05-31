<%
@ Language = VBScript
%>
<%
Option Explicit

Dim fso, testsDir, folderObj, fileObj
Dim testFiles(), testCount
Dim currentFile

Set fso = Server.CreateObject("Scripting.FileSystemObject")
' The physical path to the tests directory (where we are)
testsDir = Server.MapPath(".")

testCount = 0
ReDim testFiles(0)

If fso.FolderExists(testsDir) Then
    Set folderObj = fso.GetFolder(testsDir)
    For Each fileObj In folderObj.Files
        If LCase(Right(fileObj.Name, 4)) = ".asp" Then
            ' Only list files starting with test_
            If LCase(Left(fileObj.Name, 5)) = "test_" Then
                If LCase(fileObj.Name) <> "default.asp" Then
                    If testCount > 0 Then
                        ReDim Preserve testFiles(testCount)
                    End If
                    testFiles(testCount) = fileObj.Name
                    testCount = testCount + 1
                End If
            End If
        End If
    Next
End If

' Sort alphabetically
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
<html lang="en">

    <head>
        <meta charset="UTF-8" />
        <title>AxonASP Tests Dashboard - G3Pix</title>
        <link rel="stylesheet" href="../css/axonasp.css" />
        <style>
            /* Custom overrides for the test dashboard */
            #main-container {
                height: calc(100vh - 85px);
                /* Header + Status Bar + padding */
                overflow: hidden;
            }

            #sidebar {
                overflow-y: auto;
            }

            #content {
                padding: 0;
                overflow: hidden;
                display: flex;
                flex-direction: column;
            }

            #test-viewer-header {
                background-color: #00088f;
                border-bottom: 1px solid #808080;
                padding: 5px 15px;
                font-weight: bold;
                font-size: 11px;
                display: flex;
                justify-content: space-between;
                align-items: center;
            }

            #test-frame {
                flex: 1;
                width: 100%;
                border: none;
                background-color: #fff;
            }

            .treeview li {
                margin: 2px 0;
            }

            .treeview a.active {
                background-color: #335ea8;
                color: #fff !important;
            }
        </style>
    </head>

    <body>
        <div id="header">
            <div class="logo">
                <img src="../logo_square.svg" alt="AxonASP" width="43" />
            </div>
            <h1>AxonASP Tests Dashboard</h1>
        </div>

        <div id="main-container">
            <div id="sidebar">
                <div class="section-title">Available Tests</div>
                <ul class="treeview">
                    <%
                    If testCount = 0 Then
                    %>
                    <li>No tests found.</li>
                    <%
                    Else
                        Dim index
                        For index = 0 To testCount - 1
                            Dim fileName, activeClass
                            fileName = testFiles(index)
                            activeClass = ""
                            If LCase(fileName) = LCase(currentFile) Then
                                activeClass = " active"
                            End If
                    %>
                    <li class="file">
                        <a href="?file=<%= Server.URLEncode(fileName) %>" class="test-link<%= activeClass %>"
                            data-file="<%= fileName %>" onclick="return loadTest(this);"><%= fileName %></a>
                    </li>
                    <%
                        Next
                    End If
                    %>
                </ul>

                <div class="section-title" style="margin-top: 20px">
                    Navigation
                </div>
                <ul>
                    <li><a href="../default.asp">&laquo; Back to Home</a></li>
                    <li><a href="../manual/default.asp">Software Manual</a></li>
                </ul>
            </div>

            <div id="content">
                <div id="test-viewer-header">
                    <span id="current-filename"><%= currentFile %></span>
                    <button class="btn" onclick="refreshTest()" style="font-size: 10px; padding: 1px 5px">
                        Refresh
                    </button>
                </div>
                <%
                If currentFile = "" Then
                %>
                <div style="padding: 20px">
                    <div class="alert alert-info">
                        <strong>Information:</strong> Please select a test file
                        from the sidebar to begin execution.
                    </div>
                </div>
                <%
                Else
                %>
                <iframe id="test-frame" src="<%= currentFile %>" title="Test Viewer"></iframe>
                <%
                End If
                %>
            </div>
        </div>

        <div id="status-bar">
            <span class="status-badge"></span>
            <strong>Status:</strong> Ready | Total Tests:
            <%= testCount %>
            | Session:
            <%= Session.SessionID %>
        </div>

        <script>
            function loadTest(element) {
                var fileName = element.getAttribute("data-file");
                var frame = document.getElementById("test-frame");
                var label = document.getElementById("current-filename");

                if (frame) {
                    frame.src = fileName;
                }

                if (label) {
                    label.innerText = fileName;
                }

                // Update active state
                var links = document.querySelectorAll(".test-link");
                links.forEach(function (link) {
                    link.classList.remove("active");
                });
                element.classList.add("active");

                // Update URL without reload
                if (window.history && window.history.replaceState) {
                    window.history.replaceState(
                        null,
                        null,
                        "?file=" + encodeURIComponent(fileName)
                    );
                }

                return false;
            }

            function refreshTest() {
                var frame = document.getElementById("test-frame");
                if (frame) {
                    frame.contentWindow.location.reload();
                }
            }
        </script>
    </body>

</html>