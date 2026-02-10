
<%Class cls_fontTypeList
Public list
Private Sub Class_Initialize
Set list = Server.CreateObject("scripting.dictionary")
list.Add "Arial","Arial"
list.Add "Arial Black","Arial Black"
list.Add "Arial Narrow","Arial Narrow"
list.Add "Book Antiqua","Book Antiqua"
list.Add "Bookman Old Style","Bookman Old Style"
list.Add "Century Gothic","Century Gothic"
list.Add "Comic Sans MS","Comic Sans MS"
list.Add "Courier New","Courier New"
list.Add "Garamond","Garamond"
list.Add "Georgia","Georgia"
list.Add "Helvetica","Helvetica"
list.Add "Impact","Impact"
list.Add "Lucida Console","Lucida Console"
list.Add "Lucida Sans","Lucida Sans"
list.Add "Lucida Unicode","Lucida Unicode"
list.Add "Monotype Corsiva","Monotype Corsiva"
list.Add "Tahoma","Tahoma"
list.Add "Times New Roman","Times New Roman"
list.Add "Trebuchet","Trebuchet"
list.Add "Trebuchet MS","Trebuchet MS"
list.Add "Verdana","Verdana"
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
showSelected = showSelected & "<option value='" & key & "'"
If convertStr(selected) = convertStr(key) Then
showSelected = showSelected & " selected"
End If
showSelected = showSelected & ">" & list(key) & "</option>"
Next
End Select
End Function
End class%>
