
<%Const sb_equals="="
Const sb_lower="<"
Const sb_lowerOrEquals="<="
Const sb_greater=">"
Const sb_greaterOrEquals=">="
Class cls_lowerGreaterThanList
Public list
Private Sub Class_Initialize
Set list = Server.CreateObject("scripting.dictionary")
list.Add sb_equals, sb_equals
list.Add sb_lower, sb_lower
list.Add sb_lowerOrEquals, sb_lowerOrEquals
list.Add sb_greater, sb_greater
list.Add sb_greaterOrEquals, sb_greaterOrEquals
End Sub
Private Sub Class_Terminate
Set list = nothing
End Sub
Public Function showSelected(mode, selected)
showSelected = ""
selected=convertStr(selected)
Select Case mode
Case "single"
showSelected = list(selected)
Case "option"
Dim key
For each key in list
showSelected = showSelected & "<option value='" & key & "'"
If selected = key Then
showSelected = showSelected & " selected"
End If
showSelected = showSelected & ">" & list(key) & "</option>"
Next
End Select
End Function
End class%>
