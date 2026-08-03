<%@ Language="VBScript" %>
<%
Response.ContentType = "text/plain"

On Error Resume Next

Dim targetFolderPath, filePath, fso, stream, errNum, errDesc, folderExists

targetFolderPath = Server.MapPath("./temp_delete_test_folder")

Set fso = Server.CreateObject("Scripting.FileSystemObject")

If fso.FolderExists(targetFolderPath) Then
    fso.DeleteFolder targetFolderPath, True
    Err.Clear
End If

fso.CreateFolder targetFolderPath

filePath = fso.BuildPath(targetFolderPath, "childFile.txt")
Set stream = fso.CreateTextFile(filePath, True)
stream.Write "child file content"
stream.Close
Set stream = Nothing

fso.DeleteFolder targetFolderPath, True
errNum = Err.Number
errDesc = Err.Description
Err.Clear

folderExists = fso.FolderExists(targetFolderPath)

Response.Write "Target Folder Path: " & targetFolderPath & vbCrLf
Response.Write "Error Number: " & errNum & vbCrLf
Response.Write "Error Description: " & errDesc & vbCrLf
Response.Write "Folder Exists Post-Delete: " & folderExists & vbCrLf

If errNum = 0 And Not folderExists Then
    Response.Write "PASS" & vbCrLf
Else
    Response.Write "FAIL" & vbCrLf
End If
%>
