<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include virtual="/asplite/asplite.asp"-->
<%
Response.ContentType = "text/html"
Response.Buffer = False

Response.Write "<h2>Uploader with Debug Points</h2>"
Response.Flush

If Request.TotalBytes < 1 Then
    Response.Write "<p>No POST data.</p>"
    Response.End
End If

Response.Write "<p>TotalBytes: " & Request.TotalBytes & "</p>"
Response.Flush

' Load original uploader and add debug points
Dim objStream, fileContent
Set objStream = Server.CreateObject("ADODB.Stream")
objStream.CharSet = "utf-8"
objStream.Open
objStream.Type = 2
objStream.LoadFromFile Server.MapPath("/asplite/plugins/uploader/uploader.asp")
fileContent = objStream.ReadText()
objStream.Close
Set objStream = Nothing

Response.Write "<p>Original file: " & Len(fileContent) & " chars</p>"
Response.Flush

' Remove code blocks
fileContent = Replace(fileContent, "<%", "", 1, -1, 1)
fileContent = Replace(fileContent, "%>", "", 1, -1, 1)

' Add debug at key points
' After Class_Initialize start
fileContent = Replace(fileContent, _
    "Private Sub Class_Initialize()", _
    "Private Sub Class_Initialize()" & vbCrLf & "Response.Write ""<p>DEBUG: Class_Initialize start</p>"" : Response.Flush")

' After setting dictionaries
fileContent = Replace(fileContent, _
    "Set FormElements = aspL.dict", _
    "Set FormElements = aspL.dict" & vbCrLf & "Response.Write ""<p>DEBUG: Dictionaries set</p>"" : Response.Flush")

' After stream open
fileContent = Replace(fileContent, _
    "StreamRequest.Open", _
    "StreamRequest.Open" & vbCrLf & "Response.Write ""<p>DEBUG: Stream opened</p>"" : Response.Flush")

' At Upload() start
fileContent = Replace(fileContent, _
    "Public Sub Upload()", _
    "Public Sub Upload()" & vbCrLf & "Response.Write ""<p>DEBUG: Upload() start</p>"" : Response.Flush")

' After uploadedYet = true
fileContent = Replace(fileContent, _
    "uploadedYet = true", _
    "uploadedYet = true" & vbCrLf & "Response.Write ""<p>DEBUG: uploadedYet set</p>"" : Response.Flush")

' Before first BinaryRead
fileContent = Replace(fileContent, _
    "readBytes = internalChunkSize", _
    "Response.Write ""<p>DEBUG: Before readBytes = internalChunkSize</p>"" : Response.Flush" & vbCrLf & "readBytes = internalChunkSize")

' After first BinaryRead
fileContent = Replace(fileContent, _
    "VarArrayBinRequest = Request.BinaryRead(readBytes)", _
    "VarArrayBinRequest = Request.BinaryRead(readBytes)" & vbCrLf & "Response.Write ""<p>DEBUG: After first BinaryRead, readBytes="" & readBytes & ""</p>"" : Response.Flush")

' Before the loop
fileContent = Replace(fileContent, _
    "Do Until readBytes < 1", _
    "Response.Write ""<p>DEBUG: Entering loop</p>"" : Response.Flush" & vbCrLf & "Do Until readBytes < 1")

' Inside the loop
fileContent = Replace(fileContent, _
    "tmpBinRequest = Request.BinaryRead(readBytes)", _
    "Response.Write ""<p>DEBUG: Loop iteration, readBytes="" & readBytes & ""</p>"" : Response.Flush" & vbCrLf & "tmpBinRequest = Request.BinaryRead(readBytes)" & vbCrLf & "Response.Write ""<p>DEBUG: After loop read, readBytes="" & readBytes & ""</p>"" : Response.Flush")

Response.Write "<p>Modified file: " & Len(fileContent) & " chars</p>"
Response.Flush

Response.Write "<hr><p>Executing code...</p>"
Response.Flush

On Error Resume Next
ExecuteGlobal fileContent
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ExecuteGlobal error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
End If
Response.Flush

Response.Write "<hr><p>Creating instance...</p>"
Response.Flush

Dim upload
Set upload = Eval("new cls_asplite_uploader")
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>New error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
End If
Response.Flush

Response.Write "<hr><p>Calling Upload()...</p>"
Response.Flush

upload.Upload
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Upload() error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
End If
Response.Flush

Response.Write "<p>UploadedFiles.Count: " & upload.UploadedFiles.Count & "</p>"
On Error GoTo 0

Response.Write "<hr><p><b>Test complete!</b></p>"
%>
