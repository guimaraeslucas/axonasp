<%
Dim arr, result
arr = Array("Hello World", "Goodbye")
result = Filter(arr, "Hello World")
Response.Write "Result type: " & TypeName(result) & "<br>"
Response.Write "Result(0): " & result(0) & "<br>"
%>
