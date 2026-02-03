<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include virtual="/asplite/asplite.asp"-->
<%
Response.ContentType = "text/html"
Response.Buffer = False

Response.Write "<h2>aspLite Plugin Load Simulation</h2>"
Response.Flush

If Request.TotalBytes < 1 Then
    Response.Write "<p>No POST data.</p>"
    Response.End
End If

Response.Write "<p>TotalBytes: " & Request.TotalBytes & "</p>"
Response.Flush

' Simulate exactly what aspL.plugin("uploader") does
Dim pluginPath
pluginPath = "/asplite/plugins/uploader/uploader.asp"

Response.Write "<p>Loading plugin from: " & pluginPath & "</p>"
Response.Flush

' Load file exactly as aspLite does (using aspL.stream and aspL.removeCRB)
Response.Write "<p>Reading file content...</p>"
Response.Flush

Dim objStream, fileContent
Set objStream = Server.CreateObject("ADODB.Stream")
objStream.CharSet = "utf-8"
objStream.Open
objStream.Type = 2 'adTypeText
objStream.LoadFromFile Server.MapPath(pluginPath)
fileContent = objStream.ReadText()
objStream.Close
Set objStream = Nothing

Response.Write "<p>Read " & Len(fileContent) & " characters</p>"
Response.Flush

' Remove code render blocks
fileContent = Replace(fileContent, "<%", "", 1, -1, 1)
fileContent = Replace(fileContent, "%>", "", 1, -1, 1)

Response.Write "<p>After removeCRB: " & Len(fileContent) & " characters</p>"
Response.Flush

' ExecuteGlobal
Response.Write "<p>Calling ExecuteGlobal...</p>"
Response.Flush

On Error Resume Next
ExecuteGlobal fileContent
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ExecuteGlobal error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
Else
    Response.Write "<p style='color:green'>ExecuteGlobal succeeded!</p>"
End If
Response.Flush

' Create instance via Eval
Response.Write "<hr><p>Creating instance via Eval...</p>"
Response.Flush

Dim upload
Set upload = Eval("new cls_asplite_uploader")
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Eval error: " & Err.Number & " - " & Err.Description & "</p>"
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
