
<%class cls_customerList
public showRecentOnly
Private Sub Class_Initialize
showRecentOnly=false
end sub 
Public Function table
set table=server.CreateObject ("scripting.dictionary")
dim sql, rs
if showRecentOnly then
set rs=db.execute("select iId from tblCustomer order by dCreatedTS desc")
else
set rs=db.execute("select iId from tblCustomer order by sName")
end if
dim customerObj, runner,contii
runner=0
contii=true
while not rs.eof and contii
set customerObj=new cls_customer
customerObj.pick(rs(0))
table.Add customerObj.iId, customerObj
set customerObj=nothing
rs.movenext
runner=runner+1
if showRecentOnly then
if runner=20 then contii=false
end if
wend
set rs=nothing
end function
Public function getAll
set getAll=db.execute("select iId,sName,sUrl,dCreatedTS,sUrl,webmasterEMAIL,iFolderSize from tblCustomer order by dCreatedTS desc")
end function
Public Property Get count
dim rs
set rs=db.execute("select count(*) from tblCustomer")
count=convertGetal(rs(0))
set rs=nothing
end Property
end class%>
