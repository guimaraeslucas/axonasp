<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include virtual="/asplite/asplite.asp"-->
<!-- #include virtual="/tests/uploader_debug.asp"-->
<%
Response.ContentType = "text/html"
Response.Buffer = False

Response.Write "<h2>DEBUG Uploader Test</h2>"
Response.Flush

If Request.TotalBytes < 1 Then
    Response.Write "<p>No POST data. Use curl to upload.</p>"
    Response.End
End If

Response.Write "<p>TotalBytes: " & Request.TotalBytes & "</p>"
Response.Flush

' Create debug uploader
Response.Write "<hr><p><b>Creating debug uploader...</b></p>"
Response.Flush

On Error Resume Next
Dim upload
Set upload = New cls_asplite_uploader_debug

If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
End If
Response.Flush

' Call Upload
Response.Write "<hr><p><b>Calling Upload()...</b></p>"
Response.Flush

upload.Upload

If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Upload error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
End If
Response.Flush

Response.Write "<hr><p>UploadedFiles.Count: " & upload.UploadedFiles.Count & "</p>"
Response.Flush

On Error GoTo 0
Set upload = Nothing

Response.Write "<hr><p><b>Test complete!</b></p>"
%>
