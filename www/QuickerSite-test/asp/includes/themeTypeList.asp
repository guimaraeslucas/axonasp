
<%const QS_theme_cd=10
const QS_theme_pb=20
const QS_theme_ts=30
Class cls_themeTypeList
Public list
Private Sub Class_Initialize
Set list = Server.CreateObject("scripting.dictionary")
list.Add QS_theme_cd,l("communitydriven")
list.Add QS_theme_pb,l("personalblog")
list.Add QS_theme_ts,l("privatemessagesystem")
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
showSelected = showSelected & "<option value='" & key &"'"
If selected = key Then
showSelected = showSelected & " selected"
End If
showSelected = showSelected & ">" & list(key) & "</option>"
Next
End Select
End Function
End class%>
