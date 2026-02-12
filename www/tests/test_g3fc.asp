<%
' Test G3FC implementation
Response.Write "<h1>G3FC Library Test</h1>"

Set fso = Server.CreateObject("Scripting.FileSystemObject")
Set fc = Server.CreateObject("G3FC")

' 1. Prepare test environment
srcDir = Server.MapPath("temp/g3fc_src")
destDir = Server.MapPath("temp/g3fc_dest")
archivePath = "temp/test_archive.g3fc"

If Not fso.FolderExists(srcDir) Then fso.CreateFolder(srcDir)
If Not fso.FolderExists(destDir) Then fso.CreateFolder(destDir)

' Create some dummy files
fso.CreateTextFile(srcDir & "/hello.txt", True).WriteLine("Hello from G3FC!")
fso.CreateTextFile(srcDir & "/notes.txt", True).WriteLine("This is a test file.")

Response.Write "<h3>1. Creating Archive</h3>"
' Test creating with a password and some options
Set opts = Server.CreateObject("Scripting.Dictionary")
opts.Add "CompressionLevel", 3
opts.Add "GlobalCompression", True

' We can call methods directly
success = fc.Create(archivePath, "temp/g3fc_src", "p@ssword", opts)

If success Then
    Response.Write "<span style='color:green;'>Archive created successfully: " & archivePath & "</span><br>"
Else
    Response.Write "<span style='color:red;'>Failed to create archive.</span><br>"
End If

Response.Write "<h3>2. Listing Archive Contents</h3>"
files = fc.List(archivePath, "p@ssword", "KB", True)

If Not IsEmpty(files) Then
    Response.Write "<table border='1' cellpadding='5' style='border-collapse:collapse;'>"
    Response.Write "<tr><th>Path</th><th>Size</th><th>Type</th><th>Checksum</th></tr>"
    For Each f In files
        Response.Write "<tr>"
        Response.Write "<td>" & f.Item("Path") & "</td>"
        Response.Write "<td>" & f.Item("FormattedSize") & "</td>"
        Response.Write "<td>" & f.Item("Type") & "</td>"
        Response.Write "<td>" & f.Item("Checksum") & "</td>"
        Response.Write "</tr>"
    Next
    Response.Write "</table>"
Else
    Response.Write "<span style='color:red;'>Failed to list files or archive is empty.</span><br>"
End If

Response.Write "<h3>3. Extracting Archive</h3>"
' Extract to the destination directory
success = fc.Extract(archivePath, "temp/g3fc_dest", "p@ssword")

If success Then
    Response.Write "<span style='color:green;'>Archive extracted successfully to temp/g3fc_dest</span><br>"
    
    ' Verify file existence
    If fso.FileExists(destDir & "/g3fc_src/hello.txt") Then
        Response.Write "Verification: hello.txt found in destination.<br>"
        Response.Write "Content: <b>" & fso.OpenTextFile(destDir & "/g3fc_src/hello.txt").ReadAll() & "</b><br>"
    End If
Else
    Response.Write "<span style='color:red;'>Extraction failed.</span><br>"
End If

Response.Write "<h3>4. Finding Files</h3>"
found = fc.Find(archivePath, "hello", "p@ssword")
If Not IsEmpty(found) Then
    Response.Write "Search for 'hello' found " & (UBound(found) + 1) & " item(s).<br>"
    For Each item In found
        Response.Write "- Found: " & item.Item("Path") & " (" & item.Item("Size") & " bytes)<br>"
    Next
End If

Response.Write "<br><hr>Test Complete."
%>
