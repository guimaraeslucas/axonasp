<%
Option Explicit
Response.ContentType = "text/plain"

' Simple MD5 test
Dim md5Obj
Set md5Obj = aspL.plugin("md5")

Dim testInput
testInput = "password"

Response.Write "Testing MD5 with input: " & testInput & vbCrLf

Dim result
result = md5Obj.md5(testInput, 32)

Response.Write "MD5 Result: " & result & vbCrLf
Response.Write "Result Length: " & Len(result) & vbCrLf

If Len(result) > 0 Then
    Response.Write "SUCCESS - Hash generated" & vbCrLf
Else
    Response.Write "FAIL - Empty result" & vbCrLf
End If
%>
