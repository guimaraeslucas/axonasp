
<%const QS_theme_sublevel_none	= 10
const QS_theme_sublevel_authortopic	= 15
const QS_theme_sublevel_topic	= 20
const QS_theme_sublevel_theme	= 30
Class cls_theme_sublevelList
Public list
Private Sub Class_Initialize
Set list = Server.CreateObject("scripting.dictionary")
list.Add QS_theme_sublevel_none, l("sublevel_none")
list.Add QS_theme_sublevel_authortopic, l("sublevel_authortopic")
list.Add QS_theme_sublevel_topic, l("sublevel_topic")
list.Add QS_theme_sublevel_theme, l("sublevel_theme")
End Sub
Private Sub Class_Terminate
Set list = nothing
End Sub
Public Function showSelected(mode, selected)
showSelected = ""
selected=convertGetal(selected)
Select Case mode
Case "single"
showSelected = list(selected)
Case "option"
Dim key
For each key in list
showSelected = showSelected & "<option value=" & key
If selected = key Then
showSelected = showSelected & " selected"
End If
showSelected = showSelected & ">" & list(key) & "</option>"
Next
End Select
End Function
End class%>
