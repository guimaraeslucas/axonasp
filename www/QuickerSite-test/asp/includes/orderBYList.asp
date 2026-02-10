
<%Const sb_sTitle="sTitle"
Const sb_updatedTS_ASC="updatedTS ASC"
Const sb_updatedTS_DESC="updatedTS DESC"
Const sb_createdTS_ASC="createdTS ASC"
Const sb_createdTS_DESC="createdTS DESC"
Const sb_dPageDESC="dPage DESC"
Const sb_dPageASC="dPage ASC"
Class cls_orderByList
Public list
Private Sub Class_Initialize
Set list = Server.CreateObject("scripting.dictionary")
list.Add sb_sTitle, l("titleoflistitem")
list.Add sb_dPageDESC, l("dateofitem")
list.Add sb_dPageASC, l("dateofitem2")
list.Add sb_createdTS_DESC, l("dateofcreationitem")
list.Add sb_createdTS_ASC, l("dateofcreationitem2")
list.Add sb_updatedTS_DESC, l("dateofupdateitem")
list.Add sb_updatedTS_ASC, l("dateofupdateitem2")
End Sub
Private Sub Class_Terminate
Set list = nothing
End Sub
Public Function showSelected(mode, selected)
showSelected = ""
Select Case mode
Case "single"
showSelected = list(selected)
Case "option"
Dim key
For each key in list
showSelected = showSelected & "<option value='" & key & "'"
If convertStr(selected) = convertStr(key) Then
showSelected = showSelected & " selected"
End If
showSelected = showSelected & ">" & list(key) & "</option>"
Next
End Select
End Function
End class%>
