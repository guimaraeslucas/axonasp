
<%const QS_dateFM_EU="10"
const QS_dateFM_US="20"
Class cls_dateFormatList
Public list
Private Sub Class_Initialize
Set list = Server.CreateObject("scripting.dictionary")
list.Add QS_dateFM_EU,"DD/MM/YYYY"
list.Add QS_dateFM_US,"MM/DD/YYYY"
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
showSelected = showSelected & "<option value='" & key &"'"
If convertGetal(selected) = convertGetal(key) Then
showSelected = showSelected & " selected"
End If
showSelected = showSelected & ">" & list(key) & "</option>"
Next
End Select
End Function
End class%>
