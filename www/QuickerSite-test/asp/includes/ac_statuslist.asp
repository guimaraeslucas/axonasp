
<%Class cls_statusList
Public list
Private Sub Class_Initialize
Set list = Server.CreateObject("scripting.dictionary")
list.Add 20, "Confirmed"
list.Add 17, "Pending"
list.Add 15, "Unavailable"
End Sub
public function getCSS(iStatus)
getCSS="background-image:url("
getCSS=getCSS& C_DIRECTORY_QUICKERSITE & "/fixedImages/Acimages/" & iStatus & ".gif"
getCSS=getCSS& ");"
end function
public function getImage(iStatus)
getImage="<img style=""margin:0px;border-style:none"" alt="""" src=""" & C_DIRECTORY_QUICKERSITE & "/fixedImages/Acimages/" & iStatus & ".gif"" />"
end function
Public Function showSelected(mode, selected)
showSelected = ""
selected=convertStr(selected)
Select Case mode
Case "single"
showSelected = list(convertGetal(selected))
Case "option"
Dim key
For each key in list
showSelected = showSelected & "<option value=""" & key & """"
If convertGetal(selected) = convertGetal(key) Then
showSelected = showSelected & " selected"
End If
showSelected = showSelected & ">" & list(key) & "</option>"
Next
End Select
End Function
End class%>
