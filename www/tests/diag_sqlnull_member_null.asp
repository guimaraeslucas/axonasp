<%
Function isLeeg(ByVal value)
  isLeeg = False
  If IsNull(value) Then
    isLeeg = True
  Else
    If IsEmpty(value) Or Trim(value) = "" Then isLeeg = True
  End If
End Function
Function sqlNull(field, value)
  If isLeeg(value) Then
    sqlNull = field & " IS NULL"
  Else
    sqlNull = field & "=" & CStr(value)
  End If
End Function
Class C
  Public iId
  Private Sub Class_Initialize
    iId = Null
  End Sub
End Class
Dim o
Set o = New C
Response.Write "direct IsNull(o.iId)=" & CStr(IsNull(o.iId)) & "<br>"
Response.Write "sqlNull(o.iId)=" & sqlNull("x", o.iId) & "<br>"
%>
