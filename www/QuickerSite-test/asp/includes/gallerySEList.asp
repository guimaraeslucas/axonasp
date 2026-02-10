
<%const QS_gallery_SE_BW=10
const QS_gallery_SE_GS=20
const QS_gallery_SE_SE=30
Class cls_gallerySEList
Public list
Private Sub Class_Initialize
Set list = Server.CreateObject("scripting.dictionary")
list.Add QS_gallery_SE_BW,"BlackWhite"
list.Add QS_gallery_SE_GS,"Grayscale"
list.Add QS_gallery_SE_SE,"Sepia"
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
If convertStr(selected) = convertStr(key) Then
showSelected = showSelected & " selected"
End If
showSelected = showSelected & ">" & list(key) & "</option>"
Next
End Select
End Function
End class%>
