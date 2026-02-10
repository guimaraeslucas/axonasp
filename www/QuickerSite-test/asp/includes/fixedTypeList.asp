
<%Const sb_item=10
Const sb_container=20
Const sb_externalUrl=30
Const sb_list=40
Const sb_lossePagina=50
Const sb_menugroup=60
Const sb_constant=70
Class cls_fixedTypeList
Public list
Private Sub Class_Initialize
Set list = Server.CreateObject("scripting.dictionary")
list.Add sb_item, l("itemwithcontent")
list.Add sb_container, l("containeritem")
list.Add sb_externalUrl, l("externalurl")
list.Add sb_list, l("list")
list.Add sb_lossePagina, l("freepage")
'list.Add sb_constant, l("constant")
'list.Add sb_menugroup,l("menugroup")
End Sub
Private Sub Class_Terminate
Set list = nothing
End Sub
Public Function showSelected(mode, selected)
showSelected = ""
if isnull(selected) or isEmpty(selected) or selected="" then selected=0 else selected=clng(selected)
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
