<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit

' Test colon-separated inline statements

Dim result, value, blockURL

' Test 1: Simple IF with colon separator
value = 5
If value = 5 Then result = 1 : result = result + 1
Response.Write "Test 1 (IF with colon): " & result & "<br />"

' Test 2: SELECT CASE with colon separators
result = 0
value = 2
Select Case value
	Case 1 : result = 10 : result = result + 1
	Case 2 : result = 20 : result = result + 2
	Case Else : result = 30 : result = result + 3
End Select
Response.Write "Test 2 (SELECT CASE with colons): " & result & "<br />"

' Test 3: Multiple colons in sequence
result = 0
value = 3
If value = 3 Then result = 100 : result = result + 10 : result = result + 5
Response.Write "Test 3 (Multiple colons): " & result & "<br />"

' Test 4: Nested IF statements with colons and End If (from aspLite example)
blockURL = "test.php?param=value"
If InStr(blockURL, "?") > 0 Then blockURL = Left(blockURL, InStr(blockURL, "?") - 1) : End If
Response.Write "Test 4 (Complex blockURL): " & blockURL & "<br />"

Response.Write "All tests completed!"
%>

