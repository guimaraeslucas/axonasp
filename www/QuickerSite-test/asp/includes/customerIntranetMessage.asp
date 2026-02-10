
<%class cls_customerImess
Public iStatus
Public bEnabled
Public sSubject
Public sBody
Private Sub Class_Initialize
bEnabled=true
end sub
Public Function Pick(status)
dim sql, RS
if isNumeriek(status) then
sql = "select * from tblCustomerIntranetMessage where iCustomerID="&cid&" and iStatus=" & status
set RS = db.execute(sql)
if not rs.eof then
iStatus	= rs("iStatus")
bEnabled	= rs("bEnabled")
sSubject	= rs("sSubject")
sBody	= rs("sBody")
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
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
rs.Open "select * from tblCustomerIntranetMessage where iCustomerID="&cId&" and iStatus="& convertGetal(iStatus)
if rs.eof then
rs.AddNew()
end if
rs("iStatus")	= iStatus
rs("bEnabled")	= bEnabled
rs("sSubject")	= sSubject
rs("iCustomerID")	= cId
rs("sBody")	= sBody
rs.Update 
rs.close
Set rs = nothing
end function
public function getRequestValues()
end function
end class%>
