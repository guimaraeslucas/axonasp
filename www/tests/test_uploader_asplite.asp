<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include virtual="/asplite/asplite.asp"-->
<%
Response.ContentType = "text/html"
Response.Buffer = False

Response.Write "<h2>aspLite Uploader Debug Test</h2>"
Response.Flush

' Check for POST data
If Request.TotalBytes < 1 Then
    Response.Write "<p>No POST data. Use curl or form to upload.</p>"
    Response.End
End If

Response.Write "<p>TotalBytes: " & Request.TotalBytes & "</p>"
Response.Flush

' Test aspL.dict first
Response.Write "<p>Testing aspL.dict...</p>"
Response.Flush
On Error Resume Next
Dim testDict
Set testDict = aspL.dict
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>aspL.dict error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
Else
    Response.Write "<p style='color:green'>aspL.dict works! Type: " & TypeName(testDict) & "</p>"
End If
On Error GoTo 0
Response.Flush

' Create the uploader plugin
Response.Write "<hr><p>Creating uploader plugin...</p>"
Response.Flush

On Error Resume Next
Err.Clear

Dim upload
Set upload = aspL.plugin("uploader")

If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Plugin error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
End If

If upload Is Nothing Then
    Response.Write "<p style='color:red'>upload is Nothing!</p>"
Else
    Response.Write "<p style='color:green'>Plugin created: " & TypeName(upload) & "</p>"
End If
Response.Flush

' Get upload path
Dim uploadPath
uploadPath = Server.MapPath("../uploads")
Response.Write "<p>Upload path: " & uploadPath & "</p>"
Response.Flush

' Call Upload explicitly first
Response.Write "<hr><p><b>Calling upload.Upload() explicitly...</b></p>"
Response.Flush

upload.Upload

If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Upload() error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
Else
    Response.Write "<p style='color:green'>Upload() completed!</p>"
End If
Response.Flush

Response.Write "<p>UploadedFiles.Count: " & upload.UploadedFiles.Count & "</p>"
Response.Flush

' List files before save
If upload.UploadedFiles.Count > 0 Then
    Response.Write "<p>Files parsed:</p><ul>"
    Dim k, fItem
    For Each k In upload.UploadedFiles.Keys
        Set fItem = upload.UploadedFiles.Item(k)
        Response.Write "<li>" & k & ": " & fItem.FileName & " (start=" & fItem.Start & ", len=" & fItem.Length & ")</li>"
    Next
    Response.Write "</ul>"
    Response.Flush
End If

' Now call Save
Response.Write "<hr><p><b>Calling upload.Save()...</b></p>"
Response.Flush

upload.Save uploadPath

If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Save() error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
Else
    Response.Write "<p style='color:green'>Save() completed!</p>"
End If
Response.Flush

' Check error message
If upload.errorMessage <> "" Then
    Response.Write "<p>Uploader error: " & upload.errorMessage & "</p>"
End If

' List files after save
If upload.UploadedFiles.Count > 0 Then
    Response.Write "<p>Files after save:</p><ul>"
    For Each k In upload.UploadedFiles.Keys
        Set fItem = upload.UploadedFiles.Item(k)
        Response.Write "<li>" & fItem.FileName & " → " & fItem.Path & "</li>"
    Next
    Response.Write "</ul>"
End If

Set upload = Nothing
On Error GoTo 0

Response.Write "<hr><p><b>Test complete!</b></p>"
%>
