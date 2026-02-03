<%@ Language="VBScript" %>
<%
Response.ContentType = "text/html"
Response.Buffer = False

Response.Write "<h2>WriteText vs Write Test</h2>"

If Request.TotalBytes < 1 Then
    Response.Write "<p>No data. POST something to test.</p>"
    Response.End
End If

Response.Write "<p>TotalBytes: " & Request.TotalBytes & "</p>"

' Read data
Dim data, readBytes
readBytes = 200000
data = Request.BinaryRead(readBytes)
Response.Write "<p>Read " & LenB(data) & " bytes (readBytes=" & readBytes & ")</p>"
data = MidB(data, 1, LenB(data))
Response.Write "<p>After MidB: " & LenB(data) & " bytes</p>"

' Test Write
Response.Write "<hr><p><b>Testing Stream.Write (binary)...</b></p>"
Dim stream1
Set stream1 = Server.CreateObject("ADODB.Stream")
stream1.Type = 1
stream1.Open
Response.Write "<p>Stream1 opened</p>"

On Error Resume Next
stream1.Write data
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Write Error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
Else
    Response.Write "<p style='color:green'>Write succeeded! Size: " & stream1.Size & "</p>"
End If
stream1.Close
Set stream1 = Nothing

' Test WriteText
Response.Write "<hr><p><b>Testing Stream.WriteText...</b></p>"
Dim stream2
Set stream2 = Server.CreateObject("ADODB.Stream")
stream2.Type = 1
stream2.Open
Response.Write "<p>Stream2 opened (Type=1 binary)</p>"

Err.Clear
stream2.WriteText data
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>WriteText Error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
Else
    Response.Write "<p style='color:green'>WriteText succeeded! Size: " & stream2.Size & "</p>"
End If
Response.Write "<p>Calling Flush...</p>"
stream2.Flush
Response.Write "<p>Flush complete</p>"
stream2.Close
Set stream2 = Nothing

On Error GoTo 0

Response.Write "<hr><p>Test complete!</p>"
%>
