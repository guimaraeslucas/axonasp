
<%class cls_newsletterSubscriber
Public iId,sName,sEmail
Private Sub Class_Initialize
On Error Resume Next
pick(decrypt(request("iSubscriberID")))
On Error Goto 0
end sub
Public Function getRequestValues()
sName=left(request.form("sName"),50)
sEmail=left(request.form("sEmail"),50)
end Function
Public Function Pick(id)
dim sql, rs
if isNumeriek(id) then
sql = "select * from tblNewsletterCategorySubscriber where iId=" & id
set rs = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
sName	= rs("sName")
sEmail	= rs("sEmail")
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
if isLeeg(sName) then
check=false
message.AddError("err_mandatory")
end if
End Function
Public Function Save()
if check() then
Save=true
else
Save=false
exit function
end if
'set db=nothing
'set db=new cls_database
dim rs
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblNewsletterCategory where 1=2"
rs.AddNew
else
rs.Open "select * from tblNewsletterCategory where iId="& iId
end if
rs("sName")	= sName
rs("iCustomerID")	= cId
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
end function
public function remove
if not isLeeg(iId) then
dim rs
set rs=db.execute("delete from tblNewsletterCategorySubscriber where iCategoryID=" & iId)
set rs=nothing
set rs=db.execute("delete from tblNewsletterCategory where iId=" & iId)
set rs=nothing
end if
end function
Public property get nmbrSubscribers
dim rs
set rs=db.execute("select count(*) from tblNewsletterCategorySubscriber where iCategoryID=" & iId)
nmbrSubscribers=convertgetal(rs(0))
set rs=nothing
end property
end class%>
