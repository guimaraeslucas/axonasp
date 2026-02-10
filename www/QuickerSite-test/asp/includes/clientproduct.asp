
<%class cls_clientproduct
public iId,sName,dCreatedTS,dUpdatedTS,iPrice,iVat,iRenewal,dLastRenewalDate,sNotes,iClientID,dStartService,dEndService,sInvoiceNr
private sub class_initialize
On Error Resume Next
iClientID=decrypt(request("iClientID"))
dStartService=date()
dEndService=dateAdd("yyyy",1,date())
iRenewal=12
pick(decrypt(request("iClientProductId")))
On Error Goto 0
end sub
private sub class_terminate
end sub
public function client
set client=new cls_client
client.pick(iClientID)
end function
private function check
check=true
if isLeeg(sName) then
message.Adderror("err_mandatory")
check=false
end if
if convertGetal(iClientID)=0 then
message.Adderror("err_mandatory")
check=false
end if
end function
public function pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblClientProduct where iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
sName	= rs("sName")
iPrice	= rs("iPrice")
iVat	= rs("iVat")
iRenewal	= rs("iRenewal")
dLastRenewalDate	= rs("dLastRenewalDate")
sNotes	= rs("sNotes")
iClientID	= rs("iClientID")
dCreatedTS	= rs("dCreatedTS")
dUpdatedTS	= rs("dUpdatedTS")
dStartService	= rs("dStartService")
dEndService	= rs("dEndService")
sInvoiceNr	= rs("sInvoiceNr")
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
rs.Open "select * from tblClientProduct where 1=2"
rs.AddNew
rs("dCreatedTS")=now()
else
rs.Open "select * from tblClientProduct where iId="& iId
end if
rs("sName")	= sName
rs("iPrice")	=iPrice
rs("iVat")	=iVat
rs("iRenewal")	=iRenewal
rs("dLastRenewalDate")	=dLastRenewalDate
rs("sNotes")	=sNotes
rs("iClientID")	=iClientID
rs("dEndService")	=dEndService
rs("dStartService")	=dStartService
rs("sInvoiceNr")	=sInvoiceNr
rs("dUpdatedTS")	= now()
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
end function
public function list
on error resume next
dim rs
set rs = db.GetDynamicRS
rs.Open "select iId,sName,iRenewal,iPrice,iClientID,dLastRenewalDate,dEndService from tblClientProduct order by sName"
if not rs.eof then
list=rs.getrows()
else
list=null
end if
set rs=nothing
on error goto 0
end function
public function remove
if not isLeeg(iId) then
dim rs
set rs=db.execute("delete from tblClientProduct where iId="& iId)
set rs=nothing
end if
end function
end class%>
