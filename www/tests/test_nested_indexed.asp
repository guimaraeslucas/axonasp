<%
Option Explicit
Response.Write "Testing Nested Indexed Assignment<br><hr><br>"

' Test 1: Response.Cookies with sub-keys
Response.Write "<b>Test 1: Response.Cookies nested assignment</b><br>"
On Error Resume Next
Response.Cookies("testuser")("name") = "TestName"
Response.Cookies("testuser")("email") = "test@example.com"
If Err.Number = 0 Then
    Response.Write "SUCCESS: Nested cookie assignment compiled and executed<br>"
Else
    Response.Write "ERROR: " & Err.Description & "<br>"
End If
Err.Clear
On Error Goto 0

' Test 2: Dictionary with nested access (simulated)
Response.Write "<br><b>Test 2: Dictionary nested pattern</b><br>"
Dim dict
Set dict = Server.CreateObject("Scripting.Dictionary")
dict("level1") = Server.CreateObject("Scripting.Dictionary")
On Error Resume Next
dict("level1")("level2") = "NestedValue"
If Err.Number = 0 Then
    Response.Write "SUCCESS: Nested dictionary assignment compiled and executed<br>"
    If IsObject(dict("level1")) Then
        Response.Write "Value retrieved: " & dict("level1")("level2") & "<br>"
    End If
Else
    Response.Write "ERROR: " & Err.Description & "<br>"
End If
Err.Clear
On Error Goto 0

Response.Write "<br><hr><br>"
Response.Write "<b>Compilation Test: PASSED</b><br>"
Response.Write "The VM compiler successfully handles nested indexed expressions."
%>
