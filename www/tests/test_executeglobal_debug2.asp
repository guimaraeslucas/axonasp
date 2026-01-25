<%
Response.Write "<h1>ExecuteGlobal Test 2</h1>"

' Test: Assignment without Dim
Response.Write "<h2>Test: Direct Assignment</h2>"
ExecuteGlobal "testVar2 = 'Hello from ExecuteGlobal'"
Response.Write "testVar2 = " & testVar2 & "<br>"

%>
