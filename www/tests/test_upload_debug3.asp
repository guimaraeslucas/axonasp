<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include virtual="/asplite/asplite.asp"-->
<%
Response.ContentType = "text/plain"
Response.Write "=== UPLOADER DEBUG ===" & vbCrLf

On Error Resume Next

' Create uploader
Dim upload
Set upload = aspL.plugin("uploader")

If upload Is Nothing Then
    Response.Write "ERROR: Could not load uploader plugin!" & vbCrLf
    Response.End
End If

Response.Write "Plugin loaded." & vbCrLf

' Manually call Upload to trigger parsing
upload.Upload

Response.Write "Upload() method called." & vbCrLf

If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Description & vbCrLf
    Err.Clear
End If

' Check errorMessage from uploader
If upload.errorMessage <> "" Then
    Response.Write "Uploader Error: " & upload.errorMessage & vbCrLf
End If

' List UploadedFiles count
Response.Write "UploadedFiles.Count: " & upload.UploadedFiles.Count & vbCrLf

' List all keys
Response.Write "Keys: "
Dim k
For Each k In upload.UploadedFiles.Keys
    Response.Write k & " "
Next
Response.Write vbCrLf

' Try to save
Dim uploadPath
uploadPath = Server.MapPath("../uploads")
Response.Write "Calling Save to: " & uploadPath & vbCrLf

upload.Save uploadPath

If Err.Number <> 0 Then
    Response.Write "Save ERROR: " & Err.Description & vbCrLf
    Err.Clear
End If

' Check files again after save
Response.Write "After Save - UploadedFiles.Count: " & upload.UploadedFiles.Count & vbCrLf

For Each k In upload.UploadedFiles.Keys
    Response.Write "File: " & k & vbCrLf
    Response.Write "  FileName: " & upload.UploadedFiles(k).FileName & vbCrLf
    Response.Write "  FileType: " & upload.UploadedFiles(k).FileType & vbCrLf
    Response.Write "  Length: " & upload.UploadedFiles(k).Length & vbCrLf
    Response.Write "  Start: " & upload.UploadedFiles(k).Start & vbCrLf
    Response.Write "  Path: " & upload.UploadedFiles(k).Path & vbCrLf
Next

Set upload = Nothing

Response.Write "=== DONE ===" & vbCrLf

On Error Goto 0
%>
