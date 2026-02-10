
<%Const sb_text="10"
Const sb_textarea="30"
Const sb_richtext="32"
Const sb_date="35"
Const sb_checkbox="40"
Const sb_select="60"
Const sb_url="70"
Const sb_email="80"
Const sb_comment="90"
Class cls_fixedFieldTypeList
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
