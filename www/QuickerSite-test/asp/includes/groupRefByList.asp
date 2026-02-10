
<%Const sb_ref_REFERER="referer"
Const sb_ref_URL="url"
Const sb_ref_USERIP="userip"
Const sb_ref_LP="lp"
Const sb_ref_LA="la"
Const sb_ref_LOGFILE="fl"
Const sb_ref_QUERY="qw"
Class cls_refGroupByList
Public list
Private Sub Class_Initialize
Set list = Server.CreateObject("scripting.dictionary")
list.Add sb_ref_REFERER, l("referringurl")
list.Add sb_ref_URL, l("referringsite")
list.Add sb_ref_USERIP, "IP"
list.Add sb_ref_LA,l("language")
list.Add sb_ref_LP, l("landingpage")
list.Add sb_ref_QUERY,l("query")
list.Add sb_ref_LOGFILE,l("fullist")
End Sub
Private Sub Class_Terminate
Set list = nothing
End Sub
Public Function showSelected(mode, selected)
showSelected = ""
Select Case mode
Case "single"
showSelected = list(cstr(selected))
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
