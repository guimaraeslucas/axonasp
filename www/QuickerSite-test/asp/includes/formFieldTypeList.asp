
<%Const sb_ff_text="10"
Const sb_ff_textarea="20"
Const sb_ff_richtext="25"
Const sb_ff_date="30"
Const sb_ff_checkbox="40"
Const sb_ff_select="50"
Const sb_ff_radio="55"
Const sb_ff_url="60"
Const sb_ff_email="70"
Const sb_ff_file="80"
Const sb_ff_image="90"
Const sb_ff_comment="100"
Const sb_ff_hidden="110"
Class cls_formFieldTypeList
Public list
Private Sub Class_Initialize
Set list = Server.CreateObject("scripting.dictionary")
list.Add sb_ff_text, l("text") & " (<255 "& l("characters") &")"
list.Add sb_ff_textarea, l("textbox") &" (>255 "& l("characters") &")"
'list.Add sb_ff_richtext, "RichText"
list.Add sb_ff_date,l("date")
list.Add sb_ff_checkbox, l("Checkbox")
list.Add sb_ff_select, l("dropdown")
list.Add sb_ff_radio, l("choicelist")
list.Add sb_ff_url, l("url")
list.Add sb_ff_email, l("email")
list.Add sb_ff_file, l("file")
list.Add sb_ff_image, l("picture")
list.Add sb_ff_comment, l("comment")
if customer.bApplication then list.Add sb_ff_hidden, QS_CMS_hiddenfield
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
