<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.ContentType = "text/html"
Response.Buffer = False

Response.Write "<h2>BinaryRead via ExecuteGlobal Test</h2>"
Response.Flush

If Request.TotalBytes < 1 Then
    Response.Write "<p>No POST data.</p>"
    Response.End
End If

Response.Write "<p>TotalBytes: " & Request.TotalBytes & "</p>"
Response.Flush

' Test 1: Direct BinaryRead loop
Response.Write "<hr><p><b>Test 1: Direct BinaryRead loop</b></p>"
Response.Flush

Dim readBytes1, data1, tmp1
readBytes1 = 200000
data1 = Request.BinaryRead(readBytes1)
Response.Write "<p>First read: got " & LenB(data1) & ", readBytes=" & readBytes1 & "</p>" : Response.Flush

data1 = MidB(data1, 1, LenB(data1))
Dim loop1 : loop1 = 0
Do Until readBytes1 < 1
    loop1 = loop1 + 1
    Response.Write "<p>Loop " & loop1 & ": readBytes=" & readBytes1 & "</p>" : Response.Flush
    If loop1 > 5 Then Exit Do
    tmp1 = Request.BinaryRead(readBytes1)
    If readBytes1 > 0 Then data1 = data1 & MidB(tmp1, 1, LenB(tmp1))
Loop
Response.Write "<p>Total: " & LenB(data1) & " bytes</p>" : Response.Flush

' Reset for next test - request body is already consumed
Response.Write "<hr><p><b>Test 2: ExecuteGlobal BinaryRead</b></p>"
Response.Flush

Dim testCode
testCode = "Response.Write ""<p>Inside ExecuteGlobal</p>"" : Response.Flush" & vbCrLf
testCode = testCode & "Dim readBytes2, data2, tmp2" & vbCrLf
testCode = testCode & "readBytes2 = 200000" & vbCrLf
testCode = testCode & "Response.Write ""<p>Before BinaryRead: readBytes2="" & readBytes2 & ""</p>"" : Response.Flush" & vbCrLf
testCode = testCode & "data2 = Request.BinaryRead(readBytes2)" & vbCrLf
testCode = testCode & "Response.Write ""<p>After BinaryRead: readBytes2="" & readBytes2 & "", got="" & LenB(data2) & ""</p>"" : Response.Flush" & vbCrLf

Response.Write "<p>Executing code:</p>"
Response.Write "<pre>" & Server.HTMLEncode(testCode) & "</pre>"
Response.Flush

ExecuteGlobal testCode

Response.Write "<hr><p><b>Test complete!</b></p>"
%>
