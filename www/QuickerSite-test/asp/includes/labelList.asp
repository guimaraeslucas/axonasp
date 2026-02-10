
<%class cls_labelList
Public Function table
set table=server.CreateObject ("scripting.dictionary")
dim sql, rs
set rs=db.executeLabels("select iId from tblLabel order by sCode")
dim label
while not rs.eof
set label=new cls_label
label.pick(rs(0))
table.Add label.iId, label
set label=nothing
rs.movenext
wend
set rs=nothing
end function
Public function getRows(byref recordcount)
dim sql, rs
sql="SELECT tblLabel.sCode, tblLabelValue.sValue, tblLabel.iId FROM tblLanguage INNER JOIN (tblLabel INNER JOIN tblLabelValue ON tblLabel.iId = tblLabelValue.iLabelId) ON tblLanguage.iId = tblLabelValue.iLanguageId ORDER BY tblLabel.sCode"
set rs=db.getDynamicRSLabels
rs.open sql
if not rs.eof then
getRows	= rs.getRows()
recordcount	= rs.recordCount
else
getRows=null
end if
set rs=nothing
end function
Public function refresh(langKey)
dim sql, rs
sql="select sCode from tblLabel"
set rs=db.executeLabels(sql)
while not rs.eof
Application(QS_CMS_lb_ & langKey & rs(0))=""
rs.movenext
wend
set rs=nothing
end function
Public Function showSelected(mode, selected)
set list=table
showSelected = ""
selected=convertGetal(selected)
Select Case mode
Case "single"
showSelected = list(selected).sCode
Case "option"
Dim key
For each key in list
showSelected = showSelected & "<option value='" & key & "'"
If convertStr(selected) = convertStr(key) Then
showSelected = showSelected & " selected"
End If
showSelected = showSelected & ">" & list(key).sCode & "</option>"
Next
End Select
End Function
end class%>
