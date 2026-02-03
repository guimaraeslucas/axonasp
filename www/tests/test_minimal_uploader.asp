<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include virtual="/asplite/asplite.asp"-->
<%
Response.ContentType = "text/html"
Response.Buffer = False

Response.Write "<h2>Minimal Uploader Test</h2>"
Response.Flush

If Request.TotalBytes < 1 Then
    Response.Write "<p>No POST data.</p>"
    Response.End
End If

Response.Write "<p>TotalBytes: " & Request.TotalBytes & "</p>"
Response.Flush

' Define minimal uploader class
Dim minimalCode
minimalCode = "" & _
"Class cls_minimal_uploader" & vbCrLf & _
"  Public UploadedFiles" & vbCrLf & _
"  Private VarArrayBinRequest, StreamRequest, internalChunkSize" & vbCrLf & _
"" & vbCrLf & _
"  Private Sub Class_Initialize()" & vbCrLf & _
"    Response.Write ""<p>Class_Initialize</p>"" : Response.Flush" & vbCrLf & _
"    Set UploadedFiles = aspL.dict" & vbCrLf & _
"    Response.Write ""<p>UploadedFiles created</p>"" : Response.Flush" & vbCrLf & _
"    Set StreamRequest = Server.CreateObject(""ADODB.Stream"")" & vbCrLf & _
"    StreamRequest.Type = 2" & vbCrLf & _
"    StreamRequest.Open" & vbCrLf & _
"    Response.Write ""<p>Stream opened</p>"" : Response.Flush" & vbCrLf & _
"    internalChunkSize = 200000" & vbCrLf & _
"  End Sub" & vbCrLf & _
"" & vbCrLf & _
"  Public Sub Upload()" & vbCrLf & _
"    Response.Write ""<p>Upload() starting</p>"" : Response.Flush" & vbCrLf & _
"    Dim readBytes, tmpBinRequest" & vbCrLf & _
"    readBytes = internalChunkSize" & vbCrLf & _
"    VarArrayBinRequest = Request.BinaryRead(readBytes)" & vbCrLf & _
"    Response.Write ""<p>First read: readBytes="" & readBytes & "", got="" & LenB(VarArrayBinRequest) & ""</p>"" : Response.Flush" & vbCrLf & _
"    VarArrayBinRequest = MidB(VarArrayBinRequest, 1, LenB(VarArrayBinRequest))" & vbCrLf & _
"    Dim loopCount : loopCount = 0" & vbCrLf & _
"    Do Until readBytes < 1" & vbCrLf & _
"      loopCount = loopCount + 1" & vbCrLf & _
"      Response.Write ""<p>Loop "" & loopCount & "": readBytes="" & readBytes & ""</p>"" : Response.Flush" & vbCrLf & _
"      tmpBinRequest = Request.BinaryRead(readBytes)" & vbCrLf & _
"      Response.Write ""<p>After read: readBytes="" & readBytes & ""</p>"" : Response.Flush" & vbCrLf & _
"      If readBytes > 0 Then VarArrayBinRequest = VarArrayBinRequest & MidB(tmpBinRequest, 1, LenB(tmpBinRequest))" & vbCrLf & _
"    Loop" & vbCrLf & _
"    Response.Write ""<p>Upload() complete, total="" & LenB(VarArrayBinRequest) & ""</p>"" : Response.Flush" & vbCrLf & _
"  End Sub" & vbCrLf & _
"End Class"

Response.Write "<pre>" & Server.HTMLEncode(minimalCode) & "</pre>"
Response.Flush

Response.Write "<hr><p>Executing class definition...</p>"
Response.Flush

On Error Resume Next
ExecuteGlobal minimalCode
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ExecuteGlobal error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
End If
Response.Flush

Response.Write "<p>Creating instance...</p>"
Response.Flush

Dim upload
Set upload = New cls_minimal_uploader
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

On Error GoTo 0
Set upload = Nothing

Response.Write "<hr><p><b>Test complete!</b></p>"
%>
