
<%Class cls_contactFieldTypeList
Public list
Private Sub Class_Initialize
Set list = Server.CreateObject("scripting.dictionary")
list.Add sb_text, l("text") & " (<255 "& l("characters") &")"
list.Add sb_textarea, l("textbox") &" (>255 "& l("characters") &")"
list.Add sb_richtext, "RichText"
list.Add sb_date,l("date")
list.Add sb_checkbox, l("Checkbox")
list.Add sb_select, l("dropdown")
list.Add sb_url, l("url")
list.Add sb_email, l("email")
list.Add sb_comment, l("comment")
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
showSelected = showSelected & "<option value=" & key
If selected = key Then
showSelected = showSelected & " selected"
End If
showSelected = showSelected & ">" & list(key) & "</option>"
Next
End Select
End Function
End class%>
