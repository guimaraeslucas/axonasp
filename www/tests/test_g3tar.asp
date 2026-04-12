<%
@ CodePage = 65001
%>
<!--
	AxonASP Server - G3TAR Sample Page
	Demonstration of TAR archive functionality
	URL: http://localhost:8801/tests/test_g3tar.asp
-->
<%
Option Explicit
%>
<!DOCTYPE html>
<html>
    <head>
        <meta charset="utf-8" />
        <title>G3TAR - TAR Archive Test</title>
        <link rel="stylesheet" href="/css/axonasp.css" />
        <style>
            .section {
                margin: 20px 0;
                padding: 15px;
                border: 1px solid #999;
                background: #f9f9f9;
            }
            .code-box {
                background: #fff;
                border: 1px solid #ccc;
                padding: 10px;
                margin: 10px 0;
                font-family: monospace;
                white-space: pre-wrap;
                word-wrap: break-word;
            }
            .success {
                color: green;
                font-weight: bold;
            }
            .error {
                color: red;
                font-weight: bold;
            }
            table {
                border-collapse: collapse;
                width: 100%;
                margin: 10px 0;
            }
            th,
            td {
                border: 1px solid #999;
                padding: 8px;
                text-align: left;
            }
            th {
                background: #e8e8e8;
            }
        </style>
    </head>
    <body>
        <h1>G3TAR - TAR Archive Test</h1>
        <p>
            This page demonstrates the G3TAR native object for creating,
            listing, and extracting TAR archives.
        </p>

        <div class="section">
            <h2>Test 1: Create TAR Archive and Add Files</h2>
            <%
            Dim tar, fso, success
            Set tar = Server.CreateObject("G3TAR")
            Set fso = Server.CreateObject("Scripting.FileSystemObject")

            Dim tempDir, tarPath, testDir, testFile1, testFile2, testFile3
            tempDir = Server.MapPath("../temp")
            If Not fso.FolderExists(tempDir) Then
                fso.CreateFolder(tempDir)
            End If

            ' Create test directory structure
            testDir = Server.MapPath("../temp/tar_test")
            If Not fso.FolderExists(testDir) Then
                fso.CreateFolder(testDir)
            End If

            ' Create test files
            testFile1 = testDir & "\readme.txt"
            testFile2 = testDir & "\data.txt"
            testFile3 = testDir & "\config.txt"

            Dim tf1, tf2, tf3
            Set tf1 = fso.CreateTextFile(testFile1, True)
            tf1.Write "This is a README file for the TAR archive test."
            tf1.Close()

            Set tf2 = fso.CreateTextFile(testFile2, True)
            tf2.Write "Sample data file with some content. This file will be archived."
            tf2.Close()

            Set tf3 = fso.CreateTextFile(testFile3, True)
            tf3.Write "[Configuration]" & vbCrLf & "Setting1=Value1" & vbCrLf & "Setting2=Value2"
            tf3.Close()

            ' Create TAR archive
            tarPath = tempDir & "\archive.tar"
            Response.Write "<p><strong>Creating TAR archive:</strong> " & tarPath & "</p>"

            If tar.Create(tarPath) Then
                Response.Write "<p class='success'>✓ TAR archive created</p>"

                ' Add files
                If tar.AddFile(testFile1, "docs/readme.txt") Then
                    Response.Write "<p class='success'>✓ Added readme.txt</p>"
                Else
                    Response.Write "<p class='error'>✗ Failed to add readme.txt</p>"
                End If

                If tar.AddFile(testFile2, "docs/data.txt") Then
                    Response.Write "<p class='success'>✓ Added data.txt</p>"
                Else
                    Response.Write "<p class='error'>✗ Failed to add data.txt</p>"
                End If

                If tar.AddFile(testFile3, "config/config.txt") Then
                    Response.Write "<p class='success'>✓ Added config.txt</p>"
                Else
                    Response.Write "<p class='error'>✗ Failed to add config.txt</p>"
                End If

                ' Add text content directly
                If tar.AddText("metadata/info.txt", "Archive created on " & Now & " by AxonASP G3TAR") Then
                    Response.Write "<p class='success'>✓ Added metadata/info.txt (text)</p>"
                Else
                    Response.Write "<p class='error'>✗ Failed to add metadata/info.txt</p>"
                End If

                If tar.Close() Then
                    Response.Write "<p class='success'>✓ TAR archive finalized</p>"

                    ' Show archive size
                    Dim archiveSize
                    archiveSize = fso.GetFile(tarPath).Size
                    Response.Write "<p><strong>Archive Size:</strong> " & archiveSize & " bytes</p>"
                Else
                    Response.Write "<p class='error'>✗ Failed to finalize archive</p>"
                End If
            Else
                Response.Write "<p class='error'>✗ Failed to create TAR archive</p>"
            End If

            Set tar = Nothing
            %>
        </div>

        <div class="section">
            <h2>Test 2: List TAR Archive Contents</h2>
            <%
            Set tar = Server.CreateObject("G3TAR")

            If tar.Open(tarPath) Then
                Response.Write "<p class='success'>✓ TAR archive opened</p>"
                Response.Write "<p><strong>Total Files in Archive:</strong> " & tar.Count & "</p>"

                ' List contents
                Dim contents, i, entry
                contents = tar.List()

                Response.Write "<p><strong>Archive Contents:</strong></p>"
                Response.Write "<ol>"
                For i = 0 To UBound(contents)
                    Response.Write "<li>" & Server.HTMLEncode(contents(i)) & "</li>"
                Next
                Response.Write "</ol>"
            Else
                Response.Write "<p class='error'>✗ Failed to open TAR archive</p>"
            End If

            Set tar = Nothing
            %>
        </div>

        <div class="section">
            <h2>Test 3: Get File Information from Archive</h2>
            <%
            Set tar = Server.CreateObject("G3TAR")

            If tar.Open(tarPath) Then
                Response.Write "<p><strong>File Information:</strong></p>"
                Response.Write "<table>"
                Response.Write "<tr><th>Filename</th><th>Size</th><th>Type</th><th>Mode</th></tr>"

                Dim fileList, i, fileInfo
                fileList = tar.List()

                For i = 0 To UBound(fileList)
                    Set fileInfo = tar.GetFileInfo(fileList(i))
                    If Not fileInfo Is Nothing Then
                        Response.Write "<tr>"
                        Response.Write "<td>" & Server.HTMLEncode(fileInfo.Item("Name")) & "</td>"
                        Response.Write "<td>" & fileInfo.Item("Size") & " bytes</td>"
                        Response.Write "<td>" & fileInfo.Item("Type") & "</td>"
                        Response.Write "<td>" & fileInfo.Item("Mode") & "</td>"
                        Response.Write "</tr>"
                    End If
                Next
                Response.Write "</table>"
            Else
                Response.Write "<p class='error'>✗ Failed to open archive for reading</p>"
            End If

            Set tar = Nothing
            %>
        </div>

        <div class="section">
            <h2>Test 4: Extract TAR Archive</h2>
            <%
            Dim extractDir
            extractDir = Server.MapPath("../temp/tar_extract_" & CLng(Timer * 1000))

            Set fso = Server.CreateObject("Scripting.FileSystemObject")
            Set tar = Server.CreateObject("G3TAR")

            Response.Write "<p><strong>Extracting to:</strong> " & extractDir & "</p>"

            If tar.Open(tarPath) Then
                If tar.ExtractAll(extractDir) Then
                    Response.Write "<p class='success'>✓ Archive extracted successfully!</p>"

                    ' Verify extracted files
                    If fso.FolderExists(extractDir) Then
                        Dim extractedFolder, fileCount
                        Set extractedFolder = fso.GetFolder(extractDir)
                        fileCount = 0

                        ' Count files recursively
                        Dim fc
                        For Each fc In extractedFolder.Files
                            fileCount = fileCount + 1
                        Next

                        Response.Write "<p><strong>Files extracted:</strong> " & fileCount & "</p>"

                        ' Try to read one file
                        Dim extractedFile, readFile
                        extractedFile = extractDir & "\docs\readme.txt"
                        If fso.FileExists(extractedFile) Then
                            Set readFile = fso.OpenTextFile(extractedFile)
                            Dim fileContent
                            fileContent = readFile.ReadAll()
                            readFile.Close()

                            Response.Write "<p><strong>Sample: docs/readme.txt</strong></p>"
                            Response.Write "<div class='code-box'>" & Server.HTMLEncode(fileContent) & "</div>"
                        End If
                    End If

                    ' Cleanup
                    On Error Resume Next
                    fso.DeleteFolder(extractDir, True)
                    On Error Goto 0
                Else
                    If Len(tar.LastError) > 0 Then
                        Response.Write "<p class='error'>✗ Extraction failed: " & Server.HTMLEncode(tar.LastError) & "</p>"
                    Else
                        Response.Write "<p class='error'>✗ Extraction failed</p>"
                    End If
                End If
            Else
                Response.Write "<p class='error'>✗ Failed to open archive for extraction</p>"
            End If

            Set tar = Nothing
            %>
        </div>

        <div class="section">
            <h2>Test 5: Add Multiple Files at Once (Batch)</h2>
            <%
            Dim batchTarPath, batchDir, batchFiles(), i
            batchTarPath = tempDir & "\batch_archive.tar"
            batchDir = testDir

            Set tar = Server.CreateObject("G3TAR")

            ' Create some test files
            Dim file1, file2, file3
            file1 = batchDir & "\file1.txt"
            file2 = batchDir & "\file2.txt"
            file3 = batchDir & "\file3.txt"

            Set tf1 = fso.CreateTextFile(file1, True)
            tf1.Write "Batch file 1 content"
            tf1.Close()

            Set tf2 = fso.CreateTextFile(file2, True)
            tf2.Write "Batch file 2 content"
            tf2.Close()

            Set tf3 = fso.CreateTextFile(file3, True)
            tf3.Write "Batch file 3 content"
            tf3.Close()

            ' Create array of files to add
            ReDim batchFiles(2)
            batchFiles(0) = file1
            batchFiles(1) = file2
            batchFiles(2) = file3

            Response.Write "<p><strong>Creating batch archive with multiple files...</strong></p>"

            If tar.Create(batchTarPath) Then
                If tar.AddFiles(batchFiles, "batch/") Then
                    Response.Write "<p class='success'>✓ All files added in batch mode</p>"

                    If tar.Close() Then
                        Dim batchArchiveSize
                        batchArchiveSize = fso.GetFile(batchTarPath).Size
                        Response.Write "<p><strong>Batch Archive Size:</strong> " & batchArchiveSize & " bytes</p>"
                        Response.Write "<p class='success'>✓ Batch archive created successfully!</p>"

                        ' Cleanup
                        fso.DeleteFile(batchTarPath)
                    End If
                Else
                    Response.Write "<p class='error'>✗ Failed to add files in batch</p>"
                End If
            Else
                Response.Write "<p class='error'>✗ Failed to create batch archive</p>"
            End If

            Set tar = Nothing
            %>
        </div>

        <div class="section">
            <h2>TAR Archive Overview</h2>
            <p>G3TAR provides comprehensive TAR archiving capabilities:</p>
            <ul>
                <li>
                    <strong>Create(path)</strong> - Initialize a new TAR archive
                </li>
                <li>
                    <strong>Open(path)</strong> - Open an existing TAR archive
                    for reading
                </li>
                <li>
                    <strong>AddFile(src, nameInTar)</strong> - Add a file to the
                    archive
                </li>
                <li>
                    <strong>AddFolder(src, nameInTar)</strong> - Add entire
                    folder recursively
                </li>
                <li>
                    <strong>AddFiles(pathsArray, prefix)</strong> - Add multiple
                    files in batch
                </li>
                <li>
                    <strong>AddText(nameInTar, content)</strong> - Add text
                    content directly
                </li>
                <li><strong>List()</strong> - Get array of all file names</li>
                <li>
                    <strong>GetFileInfo(name)</strong> - Get metadata dictionary
                    for a file
                </li>
                <li>
                    <strong>ExtractAll(targetDir)</strong> - Extract all files
                </li>
                <li>
                    <strong>ExtractFile(name, targetDir)</strong> - Extract
                    single file
                </li>
                <li><strong>Close()</strong> - Finalize archive</li>
            </ul>
        </div>

        <div class="section">
            <h2>Cleanup</h2>
            <%
            ' Clean up test files and archives
            On Error Resume Next

            If tarPath <> "" And fso.FileExists(tarPath) Then
                fso.DeleteFile(tarPath)
                Response.Write "<p>Cleaned up: archive.tar</p>"
            End If

            If testDir <> "" And fso.FolderExists(testDir) Then
                fso.DeleteFolder(testDir, True)
                Response.Write "<p>Cleaned up: tar_test folder</p>"
            End If

            On Error Goto 0
            Response.Write "<p class='success'>✓ Temporary files cleaned up</p>"

            Set fso = Nothing
            %>
        </div>

        <hr />
        <small>AxonASP Server - G3TAR Archive Test &copy; 2026</small>
    </body>
</html>
