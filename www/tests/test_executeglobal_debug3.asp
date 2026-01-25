<%
Response.Write "<h1>ExecuteGlobal Test 3</h1>"

' Test: Assignment with proper quotes
Response.Write "<h2>Test: Direct Assignment with Double Quotes</h2>"
ExecuteGlobal "testVar3 = ""Hello from ExecuteGlobal"""
Response.Write "testVar3 = " & testVar3 & "<br>"

%>
