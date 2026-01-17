<%
Function removeCRB(value)
	value=replace(value,"<" & "%","",1,-1,1)
	removeCRB=replace(value,"%" & ">","",1,-1,1)
End Function

Response.Write "<h1>Test removeCRB</h1>"

Dim testCode
testCode = "<% Response.Write ""Hello"" %>"

Response.Write "<p>Original: " & Server.HTMLEncode(testCode) & "</p>"

Dim cleaned
cleaned = removeCRB(testCode)

Response.Write "<p>After removeCRB: " & Server.HTMLEncode(cleaned) & "</p>"

Response.Write "<p>Executing cleaned code...</p>"

ExecuteGlobal cleaned

Response.Write "<p>Done</p>"
%>
