
<%class cls_languageListNew
Public Function table
set table=server.CreateObject ("scripting.dictionary")
dim sql, rs
set rs=db.executeLabels("select iId from tblLanguage order by sLanguage")
dim language
while not rs.eof
set language=new cls_language
language.pick(rs(0))
table.Add language.iId, language
set language=nothing
rs.movenext
wend
set rs=nothing
end function
Public Function showSelected(mode, selected)
dim list
set list=table
showSelected = ""
selected=convertGetal(selected)
Select Case mode
Case "single"
showSelected = list(selected).sLanguage
Case "option"
Dim key
For each key in list
showSelected = showSelected & "<option value='" & key & "'"
If convertStr(selected) = convertStr(key) Then
showSelected = showSelected & " selected"
End If
showSelected = showSelected & ">" & list(key).sLanguage & "</option>"
Next
End Select
End Function
end class%>
