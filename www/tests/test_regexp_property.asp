<%
' Test RegExp object property assignment
' This test verifies that object properties can be set correctly
' even when Option Explicit is enabled (as in aspLite)

Response.Write "<h2>Testing RegExp Object Property Assignment</h2>"

Dim objRegExp
Set objRegExp = New RegExp

' Test property assignments (this was failing before the fix)
objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "test"

Response.Write "<p>Properties set successfully:</p>"
Response.Write "<ul>"
Response.Write "<li>IgnoreCase: " & objRegExp.IgnoreCase & "</li>"
Response.Write "<li>Global: " & objRegExp.Global & "</li>"
Response.Write "<li>Pattern: " & objRegExp.Pattern & "</li>"
Response.Write "</ul>"

' Test regex execution
Dim testString
testString = "This is a TEST string with TEST repeated"
Dim matches
Set matches = objRegExp.Execute(testString)

Response.Write "<p>Testing regex execution on: """ & testString & """</p>"
Response.Write "<p>Matches found: " & matches.Count & "</p>"

If matches.Count > 0 Then
	Response.Write "<ul>"
	Dim i
	For i = 0 To matches.Count - 1
		Response.Write "<li>Match " & (i+1) & ": """ & matches.Item(i).Value & """ at position " & matches.Item(i).FirstIndex & "</li>"
	Next
	Response.Write "</ul>"
End If

Response.Write "<p><strong style='color:green'>✓ All tests completed successfully!</strong></p>"

Set objRegExp = Nothing
%>
