<%@ Language="VBScript" %>
<%
Response.ContentType = "text/html"
Response.Write "<h2>Direct Uploader Class Test</h2>"
Response.Flush

' Include the uploader directly  
%>
<!-- #include virtual="/asplite/plugins/uploader/uploader.asp" -->
<%
Response.Write "<p>Uploader class included.</p>"
Response.Flush

' Check request
Response.Write "<p>TotalBytes: " & Request.TotalBytes & "</p>"
Response.Flush

If Request.TotalBytes < 1 Then
    Response.Write "<p>No data. POST multipart data to this endpoint.</p>"
    Response.End
End If

' Create uploader instance
Response.Write "<p>Creating uploader instance...</p>"
Response.Flush

Dim upload
Set upload = New cls_asplite_uploader
Response.Write "<p>Instance created: " & TypeName(upload) & "</p>"
Response.Flush

' Get upload path
Dim uploadPath
uploadPath = Server.MapPath("../uploads")
Response.Write "<p>Upload path: " & uploadPath & "</p>"
Response.Flush

' Call Save
Response.Write "<p>Calling upload.Save()...</p>"
Response.Flush

On Error Resume Next
upload.Save uploadPath

If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Error: " & Err.Number & " - " & Err.Description & "</p>"
Else
    Response.Write "<p style='color:green'>Save completed without errors!</p>"
End If
On Error GoTo 0

' Check results
Response.Write "<hr><p><b>Results:</b></p>"
Response.Flush

Response.Write "<p>UploadedFiles count: " & upload.UploadedFiles.Count & "</p>"
Response.Write "<p>FormData count: " & upload.FormData.Count & "</p>"
Response.Flush

If upload.UploadedFiles.Count > 0 Then
    Response.Write "<p>Files:</p><ul>"
    Dim key, fileItem
    For Each key In upload.UploadedFiles.Keys
        Set fileItem = upload.UploadedFiles.Item(key)
        Response.Write "<li>" & key & ": " & fileItem.FileName & " (" & fileItem.Size & " bytes)</li>"
    Next
    Response.Write "</ul>"
End If

Response.Write "<p>Test complete!</p>"
%>
