<%
Dim arr
arr = Array("Hello World", "Goodbye")
Response.Write "Array created: " & TypeName(arr) & "<br>"
Response.Write "Calling Filter...<br>"
On Error Resume Next
Dim result
result = Filter(arr, "Hello World")
Response.Write "Filter returned, Error: " & Err.Number & " - " & Err.Description & "<br>"
Response.Write "Result type: " & TypeName(result) & "<br>"
If IsArray(result) Then
    Response.Write "Is an array!<br>"
    Response.Write "UBound: " & UBound(result) & "<br>"
    Response.Write "result(0): " & result(0) & "<br>"
Else
    Response.Write "NOT an array<br>"
End If
%>
