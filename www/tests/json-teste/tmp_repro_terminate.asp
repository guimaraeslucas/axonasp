<%
Class JSONpair
  Dim i_value
  Public Property Get value
    If IsObject(i_value) Then
      Set value = i_value
    Else
      value = i_value
    End If
  End Property
  Public Property Let value(val)
    i_value = val
  End Property
  Public Property Set value(val)
    Set i_value = val
  End Property
  Private Sub Class_Terminate
    If IsObject(value) Then Set value = Nothing
  End Sub
End Class

Dim p
Set p = New JSONpair
p.value = 123
Set p = Nothing
Response.Write "ok"
%>
