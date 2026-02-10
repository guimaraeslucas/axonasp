
<%class cls_mail
Public iId,iCustomerID,dDateSent,sBody,sSubject,sBodyBGColor
Private sub Class_Initialize
on error resume next
pick (decrypt(request("iMail_id")))
on error goto 0
end sub
Public Function Pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblMail where iCustomerID="&cid&" and iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iID")
iCustomerID	= rs("iCustomerID")
dDateSent	= rs("dDateSent")
sBody	= rs("sBody")
sSubject	= rs("sSubject")
sBodyBGColor	= rs("sBodyBGColor")
end if
set RS = nothing
end if
end function
Public Function Save
dim rs
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblMail where 1=2"
rs.AddNew
rs("dDateSent")=now()
else
rs.Open "select * from tblMail where iId="& iId
end if
rs("iCustomerID")	= cId
rs("sBody")	= sBody
rs("sSubject")	= sSubject
rs("sBodyBGColor")	= sBodyBGColor
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
end function
public function delete()
if isNumeriek(iId) then
dim rs
set rs=db.execute("delete from tblMailContact where iMailId="& iId)
set rs=nothing
set rs=db.execute("delete from tblMail where iId="& iId)
set rs=nothing
end if
end function
Public property get iNumberRec 
dim counter
iNumberRec=receivers(counter)
iNumberRec=counter
end property
public function receivers(byref counter)
dim rs
set rs=db.getDynamicRS
rs.open "select sEmail from tblMailContact where iMailID="&iId &" order by sEmail asc"
if not rs.eof then
receivers	= rs.getRows()
counter	= rs.recordCount
end if
set rs=nothing
end function
end class%>
