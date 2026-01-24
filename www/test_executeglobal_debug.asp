<%
Response.Write "<h1>ExecuteGlobal Test</h1>"

' Test 1: Simple variable
Response.Write "<h2>Test 1: Simple Variable</h2>"
ExecuteGlobal "Dim testVar : testVar = 'Hello from ExecuteGlobal'"
Response.Write "testVar = " & testVar & "<br>"

' Test 2: Simple function
Response.Write "<h2>Test 2: Simple Function</h2>"
Dim code
code = "Function MyFunc(x)" & vbCrLf & _
       "MyFunc = x * 2" & vbCrLf & _
       "End Function"
ExecuteGlobal code
Response.Write "MyFunc(5) = " & MyFunc(5) & "<br>"

%>
