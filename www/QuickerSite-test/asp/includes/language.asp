
<%class cls_language
Public iId
Public sLanguage
Public bOnline
Public sCharset
Public sDirection
Public sCode
Private Sub Class_Initialize
iId=null
bOnline=false
sDirection=QS_ltr
end sub
Public Function Pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblLanguage where iId=" & id
set RS = db.executeLabels(sql)
if not rs.eof then
iId	= rs("iID")
sLanguage	= rs("sLanguage")
bOnline	= rs("bOnline")
sCharset	= rs("sCharset")
sDirection	= rs("sDirection")
sCode	= rs("sCode")
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
if isLeeg(sLanguage) then
check=false
end if
End Function
Public Function Save
dim rs
if check() then
save=true
else
save=false
exit function
end if
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblLanguage where 1=2"
rs.AddNew
else
rs.Open "select * from tblLanguage where iId="& iId
end if
rs("sLanguage")	= sLanguage
rs("bOnline")	= bOnline
rs("sCharset")	= sCharset
rs("sDirection")	= sDirection
rs("sCode")	= sCode
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
end function
public function getRequestValues()
sLanguage	= convertStr(Request.Form ("sLanguage"))
bOnline	= convertBool(Request.Form ("bOnline"))
sCharset	= convertStr(Request.Form ("sCharset"))
sDirection	= convertStr(Request.Form ("sDirection"))
sCode	= convertStr(Request.Form ("sCode"))
end function
end class%>
