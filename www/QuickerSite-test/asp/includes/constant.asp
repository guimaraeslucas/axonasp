
<%class cls_constant
Public iId,sConstant,sValue,iType,dOnlineFrom,dOnlineUntill,sParameters,sGlobal
Private Sub Class_Initialize
On Error Resume Next
iId	= null
iType	= convertGetal(Request("iType"))
pick(decrypt(request("iContentId")))
On Error Goto 0
end sub
Public Function Pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblConstant where iCustomerID="&cid&" and iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
sConstant	= rs("sConstant")
iType	= rs("iType")
sValue	= rs("sValue")
dOnlineFrom	= rs("dOnlineFrom")
dOnlineUntill	= rs("dOnlineUntill")
sParameters	= rs("sParameters")
sGlobal	= rs("sGlobal")
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
sConstant=Replace(sConstant, "[", "",1,-1,1)
sConstant=Replace(sConstant, "]", "",1,-1,1)
if isLeeg(sConstant) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(sValue) then
check=false
message.AddError("err_mandatory")
end if
dim checkRS
set checkRS=db.execute("select count(iId) from tblConstant where iId<>" & convertGetal(iId) & " and iCustomerID=" & cID & " and sConstant='" & ucase(sConstant) & "'")
if clng(checkRS(0))>0 then
check=false
message.AddError("err_constant")
end if
set checkRS=nothing
End Function
Public Function Save
if check() then
save=true
else
save=false
exit function
end if
dim rs
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblConstant where 1=2"
rs.AddNew
else
rs.Open "select * from tblConstant where iId="& iId
end if
rs("sConstant")	= ucase(left(sConstant,50))
rs("sValue")	= sValue
rs("iType")	= iType
rs("iCustomerID")	= cId
rs("dOnlineFrom")	= dOnlineFrom
rs("dOnlineUntill")	= dOnlineUntill
rs("sParameters")	= sParameters
rs("sGlobal")	= sGlobal
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
customer.cacheConstants()
end function
public function getRequestValues()
sConstant	= trim(convertStr(Request.Form ("sConstant")))
sValue	= convertStr(Request.Form ("sValue"))
iType	= convertGetal(Request.Form("iType"))
dOnlineFrom	= convertDateFromPicker(Request.Form ("dOnlineFrom"))
dOnlineUntill	= convertDateFromPicker(Request.Form ("dOnlineUntill"))
sParameters	= trim(convertStr(Request.Form ("sParameters")))
sGlobal	= convertStr(Request.Form ("sGlobal"))
end function
public function bOnline
bOnline=isBetween(dOnlineFrom,date,dOnlineUntill)
end function
Public property get statusString
if bOnline then
statusString=l("online")
else
statusString="<span style=""color:" & MYQS_offlineLinkColor & """>("&l("offline")&")</span>"
end if
end property
public function remove
if not isLeeg(iId) then
dim rs
set rs=db.execute("delete from tblConstant where iId="& iId)
set rs=nothing
customer.cacheConstants()
end if
end function
public function copy()
if isNumeriek(iId) then
iId	= null
sConstant	= l("copyof") & " " & sConstant
save()
end if
end function
end class%>
