
<%Const cs_silent=10
Const cs_profile=20
Const cs_read=30
Const cs_write=40
Const cs_readwrite=50
Class cls_contactStatusList
Public list
Private Sub Class_Initialize
Set list = Server.CreateObject("scripting.dictionary")
list.Add cs_silent, l("cs_silent")
list.Add cs_profile, l("cs_profile")
list.Add cs_read, l("cs_read")
list.Add cs_write, l("cs_write")
list.Add cs_readwrite, l("cs_readwrite")
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
showSelected = showSelected & "<option class='"
select case key
case 10
showSelected = showSelected & "cs_silent"
case 20
showSelected = showSelected & "cs_profile"
case 30
showSelected = showSelected & "cs_read"
case 40
showSelected = showSelected & "cs_write"
case 50
showSelected = showSelected & "cs_readwrite"
end select
showSelected = showSelected &"' value=" & key
If selected = key Then
showSelected = showSelected & " selected"
End If
showSelected = showSelected & ">" & list(key) & "</option>"
Next
End Select
End Function
End class%>
