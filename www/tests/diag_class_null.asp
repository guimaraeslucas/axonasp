<%
Class C
  Public iId
  Private Sub Class_Initialize
    iId = Null
  End Sub
End Class
Dim o
Set o = New C
Response.Write "IsNull=" & CStr(IsNull(o.iId)) & "<br>"
Response.Write "TypeName=" & TypeName(o.iId) & "<br>"
On Error Resume Next
Response.Write "CStr=" & CStr(o.iId) & "<br>"
Response.Write "Err=" & Err.Number & " " & Err.Description & "<br>"
Err.Clear
On Error Goto 0
%>
