
<%Const sb_cat_sTitle="sTitle"
Const sb_cat_dItemDate_ASC="dDate ASC"
Const sb_cat_dItemDate_DESC="dDate DESC"
Const sb_cat_updatedTS_ASC="dUpdatedTS ASC"
Const sb_cat_updatedTS_DESC="dUpdatedTS DESC"
Const sb_cat_createdTS_ASC="dCreatedTS ASC"
Const sb_cat_createdTS_DESC="dCreatedTS DESC"
Class cls_catalogOrderByList
Public list
Private Sub Class_Initialize
Set list = Server.CreateObject("scripting.dictionary")
list.Add sb_cat_sTitle, l("titleoflistitem")
list.Add sb_cat_dItemDate_DESC, l("dateofitem")
list.Add sb_cat_dItemDate_ASC, l("dateofitem2")
list.Add sb_cat_createdTS_DESC, l("dateofcreationitem")
list.Add sb_cat_createdTS_ASC, l("dateofcreationitem2")
list.Add sb_cat_updatedTS_DESC, l("dateofupdateitem")
list.Add sb_cat_updatedTS_ASC, l("dateofupdateitem2")
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
