<%
' Test G3ZIP implementation
Response.Write "<h1>G3ZIP Test</h1>"

Set fso = Server.CreateObject("Scripting.FileSystemObject")
testDir = Server.MapPath("temp/zip_test")
If Not fso.FolderExists(testDir) Then
    fso.CreateFolder(testDir)
End If

' Create some test files
fso.CreateTextFile(testDir & "/test1.txt", True).WriteLine("Content of file 1")
fso.CreateTextFile(testDir & "/test2.txt", True).WriteLine("Content of file 2")

zipPath = "temp/test_archive.zip"
Set zip = Server.CreateObject("G3ZIP")

Response.Write "<h3>Creating ZIP</h3>"
If zip.Create(zipPath) Then
    Response.Write "Created: " & zipPath & "<br>"
    
    If zip.AddFile("temp/zip_test/test1.txt", "file1.txt") Then
        Response.Write "Added test1.txt as file1.txt<br>"
    End If
    
    If zip.AddText("greeting.txt", "Hello World!") Then
        Response.Write "Added greeting.txt from string<br>"
    End If
    
    zip.Close()
    Response.Write "ZIP Closed.<br>"
Else
    Response.Write "Failed to create ZIP.<br>"
End If

Response.Write "<h3>Reading ZIP</h3>"
If zip.Open(zipPath) Then
    Response.Write "Opened: " & zipPath & "<br>"
    Response.Write "File Count: " & zip.Count & "<br>"
    
    files = zip.List()
    Response.Write "Files inside:<br>"
    For Each f In files
        Response.Write "- " & f & "<br>"
        
        ' Get info for one of them
        Set info = zip.GetFileInfo(f)
        If Not info Is Nothing Then
            Response.Write "  (Size: " & info.Item("Size") & " bytes, Modified: " & info.Item("Modified") & ")<br>"
        End If
    Next
    
    Response.Write "<h3>Extracting ZIP</h3>"
    extractPath = "temp/zip_extracted"
    If zip.ExtractAll(extractPath) Then
        Response.Write "Extracted to: " & extractPath & "<br>"
        
        If fso.FileExists(Server.MapPath(extractPath & "/file1.txt")) Then
            Response.Write "Verified: file1.txt exists in extraction dir.<br>"
        End If
    Else
        Response.Write "Extraction failed.<br>"
    End If
    
    zip.Close()
Else
    Response.Write "Failed to open ZIP.<br>"
End If

' Cleanup (optional)
' fso.DeleteFile(Server.MapPath(zipPath))
%>
