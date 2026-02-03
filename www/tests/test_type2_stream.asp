<%@ Language="VBScript" %>
<%
Response.ContentType = "text/html"
Response.Buffer = False

Response.Write "<h2>Stream Type 2 (Text) to Type 1 (Binary) CopyTo Test</h2>"

If Request.TotalBytes < 1 Then
    Response.Write "<p>No data. POST something to test.</p>"
    Response.End
End If

Response.Write "<p>TotalBytes: " & Request.TotalBytes & "</p>"

' Read data exactly as uploader does
Dim data, readBytes
readBytes = 200000
data = Request.BinaryRead(readBytes)
Response.Write "<p>Read " & LenB(data) & " bytes</p>"
data = MidB(data, 1, LenB(data))

' Create Type 2 (Text) stream exactly as uploader does
Response.Write "<hr><p><b>Creating Type 2 (Text) stream...</b></p>"
Dim StreamRequest
Set StreamRequest = Server.CreateObject("ADODB.Stream")
StreamRequest.Type = 2 ' adTypeText - EXACTLY as uploader does!
StreamRequest.Open
Response.Write "<p>StreamRequest created, Type=" & StreamRequest.Type & "</p>"

' WriteText exactly as uploader does
Response.Write "<p>Calling WriteText...</p>"
On Error Resume Next
StreamRequest.WriteText data
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>WriteText Error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
Else
    Response.Write "<p style='color:green'>WriteText succeeded! Size: " & StreamRequest.Size & "</p>"
End If

Response.Write "<p>Calling Flush...</p>"
StreamRequest.Flush
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Flush Error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
Else
    Response.Write "<p>Flush succeeded</p>"
End If

' Now test Position and CopyTo exactly as uploader's Save() does
Response.Write "<hr><p><b>Testing Save() logic...</b></p>"

' Create Type 1 (Binary) output stream
Dim streamFile
Set streamFile = Server.CreateObject("ADODB.Stream")
streamFile.Type = 1 ' adTypeBinary
streamFile.Open
Response.Write "<p>streamFile created, Type=" & streamFile.Type & "</p>"

' Set position (let's say start at byte 50, copy 20 bytes)
Dim testStart, testLength
testStart = 50
testLength = 20
Response.Write "<p>Setting StreamRequest.Position to " & testStart & "...</p>"
StreamRequest.Position = testStart
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Position Error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
Else
    Response.Write "<p>Position set, now: " & StreamRequest.Position & "</p>"
End If

' CopyTo
Response.Write "<p>Calling CopyTo with length " & testLength & "...</p>"
StreamRequest.CopyTo streamFile, testLength
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>CopyTo Error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
Else
    Response.Write "<p style='color:green'>CopyTo succeeded! streamFile.Size = " & streamFile.Size & "</p>"
End If

' SaveToFile
Dim testPath
testPath = Server.MapPath("../uploads/test_type2.bin")
Response.Write "<p>Saving to " & testPath & "...</p>"
streamFile.SaveToFile testPath, 2
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>SaveToFile Error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
Else
    Response.Write "<p style='color:green'>SaveToFile succeeded!</p>"
End If

On Error GoTo 0

streamFile.Close
StreamRequest.Close
Set streamFile = Nothing
Set StreamRequest = Nothing

Response.Write "<hr><p>Test complete!</p>"
%>
