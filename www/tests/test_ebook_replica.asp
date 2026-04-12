<%@ Language=VBScript %>
<%
' Simulate the include failing or aspl not being available
Dim DEBUG_ASP_CODE
DEBUG_ASP_CODE = "TRUE"

' Simulate aspl being undefined
' dim asplEvent : asplEvent=aspl.getRequest("asplEvent")

' Or let's try with an object that might not work
On Error Resume Next
dim asplEvent : asplEvent = Nothing
On Error GoTo 0

function pre(value)
	pre=replace(value,vbtab," ",1,-1,1)
	while instr(pre,"    ")<>0
		pre=replace(pre,"    ","   ",1,-1,1)
	wend 
end function
%>
<!doctype html>
<html>
<head>
<title>Test ebook.asp replica</title>
</head>
<body>
<h1>Test ebook.asp Replica</h1>

<p>Test 1: Direct pre() call</p>
<p><%= pre("test" & vbTab & "value") %></p>

<p>Test 2: With Server.HTMLEncode</p>
<pre><%= pre(Server.HTMLEncode("code" & vbTab & "sample")) %></pre>

</body>
</html>
