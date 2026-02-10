
<%Class cls_formatList
Public list
Private Sub Class_Initialize
Set list = Server.CreateObject("scripting.dictionary")
list.Add QS_textonly, l("textonly")
list.Add QS_html, "HTML"
if customer.bApplication and secondAdmin.bHomeVBScript then
list.Add QS_VBScript, "ASP/VBScript"
end if
End Sub
Private Sub Class_Terminate
Set list = nothing
End Sub
Public Function showSelected(mode, selected)
showSelected = ""
Select Case mode
Case "single"
showSelected = list(convertGetal(selected))
Case "option"
Dim key
For each key in list
showSelected = showSelected & "<option value=" & key
If convertGetal(selected) = convertGetal(key) Then
showSelected = showSelected & " selected"
End If
showSelected = showSelected & ">" & list(key) & "</option>"
Next
End Select
End Function
End class%>
