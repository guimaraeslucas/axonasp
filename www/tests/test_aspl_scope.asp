<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include virtual="/asplite/asplite.asp"-->
<%
Response.ContentType = "text/html"
Response.Buffer = False

Response.Write "<h2>aspL Availability Test</h2>"
Response.Flush

Response.Write "<p>Checking aspL...</p>"
Response.Write "<p>TypeName(aspL): " & TypeName(aspL) & "</p>"
Response.Flush

Response.Write "<p>Checking aspL.dict...</p>"
Dim testDict
Set testDict = aspL.dict
Response.Write "<p>TypeName(testDict): " & TypeName(testDict) & "</p>"
Response.Flush

Response.Write "<hr><p>Now testing ExecuteGlobal scope...</p>"

Dim testCode
testCode = "Response.Write ""<p>Inside ExecuteGlobal</p>"" : Response.Flush" & vbCrLf
testCode = testCode & "Response.Write ""<p>TypeName(aspL): "" & TypeName(aspL) & ""</p>"" : Response.Flush" & vbCrLf
testCode = testCode & "Dim td" & vbCrLf
testCode = testCode & "Set td = aspL.dict" & vbCrLf
testCode = testCode & "Response.Write ""<p>aspL.dict type: "" & TypeName(td) & ""</p>"" : Response.Flush" & vbCrLf

Response.Write "<pre>" & Server.HTMLEncode(testCode) & "</pre>"
Response.Flush

On Error Resume Next
ExecuteGlobal testCode
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
End If
On Error GoTo 0

Response.Write "<hr><p><b>Test complete!</b></p>"
%>
