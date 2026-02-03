<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include virtual="/asplite/asplite.asp"-->
<%
Response.ContentType = "text/html"
Response.Buffer = False

Response.Write "<h2>ExecuteGlobal Test for Uploader</h2>"
Response.Flush

If Request.TotalBytes < 1 Then
    Response.Write "<p>No POST data.</p>"
    Response.End
End If

Response.Write "<p>TotalBytes: " & Request.TotalBytes & "</p>"
Response.Flush

' Read the uploader file content
Response.Write "<p>Reading uploader.asp file...</p>"
Response.Flush

Dim fso, uploaderPath, uploaderFile, uploaderCode
Set fso = Server.CreateObject("Scripting.FileSystemObject")
uploaderPath = Server.MapPath("/asplite/plugins/uploader/uploader.asp")
Response.Write "<p>Path: " & uploaderPath & "</p>"
Response.Flush

Set uploaderFile = fso.OpenTextFile(uploaderPath, 1, False, -1) ' Unicode
uploaderCode = uploaderFile.ReadAll()
uploaderFile.Close
Set uploaderFile = Nothing
Set fso = Nothing

Response.Write "<p>Read " & Len(uploaderCode) & " characters</p>"
Response.Flush

' Strip <% and %> tags
uploaderCode = Replace(uploaderCode, "<%", "")
uploaderCode = Replace(uploaderCode, "%>", "")

' Execute the code
Response.Write "<p>Executing code via ExecuteGlobal...</p>"
Response.Flush

On Error Resume Next
ExecuteGlobal uploaderCode
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ExecuteGlobal error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
Else
    Response.Write "<p style='color:green'>ExecuteGlobal succeeded!</p>"
End If
Response.Flush

' Create the uploader
Response.Write "<hr><p>Creating cls_asplite_uploader...</p>"
Response.Flush

Dim upload
Set upload = New cls_asplite_uploader
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>New error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
Else
    Response.Write "<p style='color:green'>Created: " & TypeName(upload) & "</p>"
End If
Response.Flush

' Call Upload
Response.Write "<hr><p>Calling upload.Upload()...</p>"
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

On Error GoTo 0
Set upload = Nothing

Response.Write "<hr><p><b>Test complete!</b></p>"
%>
