<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include virtual="/asplite/asplite.asp"-->
<%
Response.ContentType = "text/plain"
Response.Write "Starting uploader test..." & vbCrLf
Response.Flush

' Check if we have multipart data
Dim contentType
contentType = Request.ServerVariables("HTTP_CONTENT_TYPE")
Response.Write "Content-Type: " & contentType & vbCrLf
Response.Write "TotalBytes: " & Request.TotalBytes & vbCrLf
Response.Flush

' Try to create the uploader plugin
Response.Write vbCrLf & "Creating uploader plugin..." & vbCrLf
Response.Flush

On Error Resume Next
Err.Clear

Dim upload
Set upload = aspL.plugin("uploader")

If Err.Number <> 0 Then
    Response.Write "ERROR creating plugin: " & Err.Description & vbCrLf
    Err.Clear
End If

If upload Is Nothing Then
    Response.Write "ERROR: upload object is Nothing!" & vbCrLf
Else
    Response.Write "Plugin created successfully. Type: " & TypeName(upload) & vbCrLf
End If
Response.Flush

' Try to call Save
Response.Write vbCrLf & "Calling upload.Save..." & vbCrLf
Response.Flush

Dim uploadPath
uploadPath = Server.MapPath("../uploads")
Response.Write "Upload path: " & uploadPath & vbCrLf
Response.Flush

upload.Save uploadPath

If Err.Number <> 0 Then
    Response.Write "ERROR during Save: " & Err.Description & vbCrLf
    Err.Clear
End If

Response.Write vbCrLf & "Checking uploaded files..." & vbCrLf
Response.Write "UploadedFiles.Count: " & upload.UploadedFiles.Count & vbCrLf
Response.Flush

Dim fileKey
For Each fileKey In upload.UploadedFiles.Keys
    Response.Write "File: " & fileKey & " = " & upload.UploadedFiles(fileKey).FileName & vbCrLf
    Response.Write "  Length: " & upload.UploadedFiles(fileKey).Length & vbCrLf
    ' Delete for security
    upload.UploadedFiles(fileKey).delete()
Next

Response.Write vbCrLf & "Form elements:" & vbCrLf
For Each fileKey In upload.FormElements.Keys
    Response.Write "Form: " & fileKey & " = " & upload.FormElements(fileKey) & vbCrLf
Next

If upload.errorMessage <> "" Then
    Response.Write vbCrLf & "Uploader error message: " & upload.errorMessage & vbCrLf
End If

Set upload = Nothing

On Error GoTo 0

Response.Write vbCrLf & "Test complete." & vbCrLf
%>
