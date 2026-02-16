<%@ Language=VBScript %>
<!-- Test page for legacy Scripting.FileSystemObject and File.* helpers -->
<html>
<head>
    <title>FSO Compatibility Test</title>
    <meta charset="utf-8">
    <style>body{font-family: Tahoma, Arial; padding:20px;}</style>
</head>
<body>
    <h2>FSO Compatibility Test</h2>

    <%
    ' Paths are relative to the web root
    Dim fname, copyname, existsBefore, existsAfter, readContent
    fname = "test_fso.txt"
    copyname = "test_fso_copy.txt"

    Response.Write "<h3>Using File.* helper</h3>"
    ' Ensure file does not exist
    If File.Exists(fname) Then File.Delete(fname)

    ' Write text using the legacy File API
    ok = File.WriteText(fname, "Hello from AxonASP FSO test")
    Response.Write "File.WriteText created: " & CStr(ok) & "<br>"

    existsAfter = File.Exists(fname)
    Response.Write "File.Exists after write: " & CStr(existsAfter) & "<br>"

    readContent = File.ReadText(fname)
    Response.Write "Read content (File.ReadText): " & Server.HTMLEncode(readContent) & "<br>"

    ' Copy using both File helper and FSO
    copyOk = File.Copy(fname, copyname)
    Response.Write "File.Copy created copy: " & CStr(copyOk) & "<br>"

    Response.Write "<h3>Using Server.CreateObject(""Scripting.FileSystemObject"")</h3>"
    Dim fso
    Set fso = Server.CreateObject("Scripting.FileSystemObject")

    Response.Write "FileExists (via FSO): " & CStr(fso.FileExists(fname)) & "<br>"
    Response.Write "FolderExists (via FSO, root): " & CStr(fso.FolderExists(".")) & "<br>"

    Dim info
    Set info = fso.GetFile(fname)
    If Not IsEmpty(info) Then
        ' Updated to use standard object property syntax (info.Name instead of info("Name"))
        Response.Write "GetFile Name: " & info.Name & " Size: " & info.Size & "<br>"
    Else
        Response.Write "GetFile returned empty<br>"
    End If

    ' FSO CopyFile and DeleteFile are Sub (Void), so we don't expect a return value.
    ' We check existence to verify success.
    fso.CopyFile fname, copyname
    Response.Write "CopyFile (via FSO) -> executed<br>"
    If fso.FileExists(copyname) Then
         Response.Write "Copy Success: True<br>"
    Else
         Response.Write "Copy Success: False<br>"
    End If

    fso.DeleteFile copyname
    Response.Write "Delete copy (via FSO) -> executed<br>"
    If Not fso.FileExists(copyname) Then
         Response.Write "Delete Success: True<br>"
    Else
         Response.Write "Delete Success: False<br>"
    End If

    ' Cleanup original
    Response.Write "Delete original (via File.Delete) -> " & CStr(File.Delete(fname)) & "<br>"

    %>

    <hr>
    <p>Notes: This test exercises both the `File.*` helper and `Server.CreateObject("Scripting.FileSystemObject")` thin shim. Some TextStream operations are delegated to `File.*` to maintain compatibility.</p>
</body>
</html>
