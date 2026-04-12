<%
' Force VM mode if possible, but our main.go usually handles it via AXONASP_VM env var
' or we can just rely on the experimental VM being called if the user specified it.

Dim a(2)
a(0) = "Hello"
a(1) = "World"
Response.Write a(0) & " " & a(1)

' Test multi-dimensional if supported
Dim b(1, 1)
b(0, 0) = "Multi"
b(1, 1) = "Array"
Response.Write "<br>" & b(0, 0) & " " & b(1, 1)
%>