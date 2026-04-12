<%
Dim a()
ReDim a(1, 2)
a(0, 0) = "X"
ReDim Preserve a(1, 4)
Response.Write "SUCCESS: " & a(0, 0) & " bounds: (" & UBound(a, 1) & "," & UBound(a, 2) & ")"
%>
