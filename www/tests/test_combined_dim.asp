<%
' Test with combined Dim
On Error Resume Next
Dim a(), b
ReDim a(1, 2)
a(0, 0) = "X"
Response.Write "Created a(1,2)<br>"
ReDim Preserve a(1, 4)
If Err.Number <> 0 Then
    Response.Write "ERROR in combined Dim test: " & Err.Number & "<br>"
Else
    Response.Write "Combined Dim works: " & a(0, 0) & "<br>"
End If
On Error Goto 0
%>
