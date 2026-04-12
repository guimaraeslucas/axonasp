<%
' Simple test for function return
On Error Resume Next

Function ReturnString()
    ReturnString = "hello"
End Function

Function ReturnNumber()
    ReturnNumber = 42
End Function

Dim s
s = ReturnString()
Response.Write "s = " & s & vbCrLf

Dim n
n = ReturnNumber()
Response.Write "n = " & n & vbCrLf
%>
