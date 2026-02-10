
<%class cls_label
Public iId
Public sCode
Public values
Private Sub Class_Initialize
iId=null
set values=server.CreateObject ("scripting.dictionary")
end sub
Public Function Pick(id)
dim sql, RS, sValue
if isNumeriek(id) then
sql = "select * from tblLabel where iId=" & id
set RS = db.executeLabels(sql)
if not rs.eof then
iId	= rs("iID")
sCode	= rs("sCode")
end if
set RS = nothing
'values ophalen
values.RemoveAll
sql="select iLanguageId,sValue from tblLabelValue where iLabelID="& iId
set rs=db.executeLabels(sql)
while not rs.eof
sValue=rs(1)
values.Add convertGetal(rs(0)),convertStr(sValue)
rs.movenext
wend
set rs=nothing
end if
end function
Public Function Check()
Check = true
if isLeeg(sCode) then
check=false
message.AddError("err_mandatory")
exit function
end if
dim languageList, languageTable
set languageList	= new cls_languageListNew
set languageTable	= languageList.table
dim lKey
for each lKey in languageTable
if isLeeg(values(lKey)) then
check=false
message.AddError("err_mandatory")
exit function
end if
next
set languageList=nothing
set languageTable=nothing
dim rs
set rs=db.executeLabels("select iId from tblLabel where iId<>"& convertGetal(iId) & " and sCode='"& sCode &"'")
if not rs.eof then
check=false
message.AddError("err_labelCode")
exit function
end if
set rs=nothing
End Function
Public Function Save
dim rs
if check() then
save=true
else
save=false
exit function
end if
set rs = db.GetDynamicRSLabels
if isLeeg(iId) then
rs.Open "select * from tblLabel where 1=2"
rs.AddNew
else
rs.Open "select * from tblLabel where iId="& iId
end if
rs("sCode")	= sCode
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
'save values
set rs=db.executeLabels("delete from tblLabelValue where iLabelId="&iId)
set rs=nothing
set rs = db.GetDynamicRSLabels
rs.open "select * from tblLabelValue where 1=2"
dim vKey
for each vKey in values 
rs.AddNew()
rs("iLabelId")=iId
rs("iLanguageId")=vKey
rs("sValue")=values(vKey)
rs.update()
next
rs.close()
set rs=nothing
end function
public function getRequestValues()
values.RemoveAll ()
dim fKey
for each fKey in Request.Form 
if left(fKey,3)="lk_" then
values.Add convertGetal(replace(fKey,"lk_","")),Request.Form(fKey)
end if
next
sCode	= replace(lcase(convertStr(Request.Form ("sCode")))," ","")
end function
public function remove
dim rs
set rs=db.executeLabels("delete from tblLabelValue where iLabelID="& convertGetal(iId))
set rs=nothing
set rs=db.executeLabels("delete from tblLabel where iId="& convertGetal(iId))
set rs=nothing
end function
end class%>
