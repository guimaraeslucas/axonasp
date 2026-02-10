
<%class cls_client
public iId,sName,sAddress,dCreatedTS,dUpdatedTS,sContactPerson,sMainEmail,sOtherEmail
private sub class_initialize
On Error Resume Next
pick(decrypt(request("iClientId")))
On Error Goto 0
end sub
private sub class_terminate
end sub
private function check
check=true
if isLeeg(sName) then
message.Adderror("err_mandatory")
check=false
end if
if isleeg(sContactPerson) then
message.Adderror("err_mandatory")
check=false
end if
if isleeg(sMainEmail) then
message.Adderror("err_mandatory")
check=false
end if
end function
public function pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblClient where iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
sName	= rs("sName")
sAddress	= rs("sAddress")
dCreatedTS	= rs("dCreatedTS")
dUpdatedTS	= rs("dUpdatedTS")
sContactPerson	= rs("sContactPerson")
sMainEmail	= rs("sMainEmail")
sOtherEmail	= rs("sOtherEmail")
end if
set RS = nothing
end if
end function
public function save()
if check() then
save=true
else
save=false
exit function
end if
dim rs
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblClient where 1=2"
rs.AddNew
rs("dCreatedTS")=now()
else
rs.Open "select * from tblClient where iId="& iId
end if
rs("sName")	= sName
rs("sAddress")	= sAddress
rs("sOtherEmail")	= sOtherEmail
rs("sContactPerson")	= sContactPerson
rs("sMainEmail")	= sMainEmail
rs("dUpdatedTS")	= now()
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
end function
public function listAll
dim rs
set rs = db.GetDynamicRS
rs.open "select iId,sName,sContactPerson,sMainEmail from tblClient order by sName asc"
if not rs.eof then
listAll=rs.getRows()
else
listAll=null
end if
Set rs = nothing
end function
public function showSelected(selected)
dim rs
set rs = db.GetDynamicRS
rs.open "select iId,sName from tblClient order by sName asc"
while not rs.eof
showSelected=showSelected& "<option value=""" & encrypt(rs("iId")) & """"
if convertGetal(selected)=convertGetal(rs("iId")) then
showSelected=showSelected & " selected=""selected"" "
end if
showSelected=showSelected & ">" & sanitize(rs("sName")) & "</option>"
rs.movenext
wend
Set rs = nothing
end function
public function remove
if not isLeeg(iId) then
dim rs
set rs=db.execute("delete from tblClientProduct where  iClientID="& iId)
set rs=nothing
set rs=db.execute("delete from tblClient where  iId="& iId)
set rs=nothing
end if
end function
public function products
set products=server.createobject("scripting.dictionary")
dim rs,cp
set rs=db.execute("select iId from tblClientProduct where iClientID=" & convertGetal(iId) & " order by sName")
while not rs.eof
set cp=new cls_clientproduct
cp.pick(rs(0))
products.Add cp.iId,cp
set cp=nothing
rs.movenext
wend
set rs=nothing
end function
end class%>
