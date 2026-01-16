<%
Response.Write "Start<br>"

Function Test1()
    Test1 = "OK"
End Function

Response.Write "Middle<br>"

Dim x
x = Test1()

Response.Write "Result: " & x & "<br>"
Response.Write "End<br>"
%>
