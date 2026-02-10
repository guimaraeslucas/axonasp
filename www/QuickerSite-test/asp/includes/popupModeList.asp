
<%Class cls_popupModeList
Public list
Private Sub Class_Initialize
Set list = Server.CreateObject("scripting.dictionary")
list.Add 0,"Show once/visit"
list.Add 1,"Show once/visitor"
list.Add 3,"Show after 2 pageloads"
list.Add 4,"Show after 3 pageloads"
list.Add 5,"Show after 4 pageloads"
list.Add 6,"Show after 5 pageloads"
list.Add 7,"Show after 6 pageloads"
list.Add 8,"Show after 7 pageloads"
list.Add 9,"Show after 8 pageloads"
list.Add 10,"Show after 9 pageloads"
End Sub
Private Sub Class_Terminate
Set list = nothing
End Sub
Public Function showSelected(mode, selected)
showSelected = ""
 selected=convertGetal(selected)
 
'response.write selected
Select Case mode
Case "single"
showSelected = list(selected)
Case "option"
Dim key
For each key in list
showSelected = showSelected & "<option value='" & key &"'"
If convertGetal(selected) = convertGetal(key) Then
showSelected = showSelected & " selected"
End If
showSelected = showSelected & ">" & list(key) & "</option>"
Next
End Select
End Function
End class%>
