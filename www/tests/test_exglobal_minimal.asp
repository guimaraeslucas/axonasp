<%
Response.Write "<h1>Test</h1>"

Function F1()
    F1 = "works"
End Function

ExecuteGlobal "Response.Write ""From ExecuteGlobal<br>"""

Response.Write "After ExecuteGlobal<br>"
Response.Write F1() & "<br>"
%>
