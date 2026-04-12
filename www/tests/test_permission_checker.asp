<%
Option Explicit

Response.Write "<h1>MSWC.PermissionChecker Test</h1>"

Dim pc
Set pc = Server.CreateObject("MSWC.PermissionChecker")

Dim testPath
testPath = "test_permission_checker.asp" ' This file itself

Response.Write "Checking access to: " & testPath & "<br>"
If pc.HasAccess(testPath) Then
    Response.Write "<span style='color:green'>Access Granted</span><br>"
Else
    Response.Write "<span style='color:red'>Access Denied</span><br>"
End If

testPath = "/nonexistent.txt"
Response.Write "Checking access to: " & testPath & "<br>"
If pc.HasAccess(testPath) Then
    Response.Write "<span style='color:red'>Access Granted (Error)</span><br>"
Else
    Response.Write "<span style='color:green'>Access Denied (Correct)</span><br>"
End If

Set pc = Nothing
%>
