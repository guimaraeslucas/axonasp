<%@ Language="VBScript" %>
<%
Response.ContentType = "text/plain"

On Error Resume Next

Response.Write "=== Testing ASPLite Uploader Plugin ===" & vbCrLf

' Get the uploader plugin
Dim upload
Set upload = aspL.plugin("uploader")

If upload Is Nothing Then
    Response.Write "ERROR: Could not load uploader plugin!" & vbCrLf
    Response.End
End If

Response.Write "Plugin loaded successfully." & vbCrLf

' Set upload directory
Dim uploadsDirVar
uploadsDirVar = Server.MapPath("../uploads")
Response.Write "Upload directory: " & uploadsDirVar & vbCrLf

' Try to save
upload.Save uploadsDirVar

Response.Write "Save() called." & vbCrLf

' Check for errors
If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf
    Err.Clear
End If

' List uploaded files
Dim fileKey
Response.Write "Uploaded files: " & vbCrLf
For Each fileKey In upload.UploadedFiles.Keys
    Response.Write "  Key: " & fileKey & vbCrLf
    Response.Write "  Filename: " & upload.UploadedFiles(fileKey).FileName & vbCrLf
    Response.Write "  Length: " & upload.UploadedFiles(fileKey).Length & vbCrLf
    Response.Write "  Path: " & upload.UploadedFiles(fileKey).Path & vbCrLf
Next

Response.Write "=== Done ===" & vbCrLf

Set upload = Nothing

On Error Goto 0
%>
