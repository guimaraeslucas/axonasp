
<%class cls_newsletterMailing
Public iId,bSent,iNewsletterID,dSentDate,sCategory,bLog
Private Sub Class_Initialize
On Error Resume Next
bLog=false
pick(decrypt(request("iNewsletterMailingID")))
On Error Goto 0
end sub
Public Function getRequestValues()
iNewsletterID	= convertgetal(Request.Form ("iNewsletterID"))
sCategory	= request.form("sCategory")
bLog	= convertBool(request.form("bLog"))
end Function
Public Function Pick(id)
dim sql, rs
if isNumeriek(id) then
sql = "select * from tblNewsletterMailing where iCustomerID="&cid&" and iId=" & id
set rs = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
iNewsletterID	= rs("iNewsletterID")
sCategory	= rs("sCategory")
bLog	= rs("bLog")
dSentDate	= rs("dSentDate")
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
if isLeeg(iNewsletterID) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(sCategory) then
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
rs.Open "select * from tblNewsletterMailing where 1=2"
rs.AddNew
else
rs.Open "select * from tblNewsletterMailing where iId="& iId
end if
rs("iNewsletterID")	= iNewsletterID
rs("sCategory")	= sCategory
rs("iCustomerID")	= cId
rs("bLog")	= bLog
rs("dSentDate")	= dSentDate
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
end function
Public function Newsletter
set Newsletter=new cls_newsletter
Newsletter.pick(iNewsletterID)
end function
Public function category
set category=new cls_newsletterCategory
category.pick(sCategory)
end function
public function remove
if not isLeeg(iId) then
dim rs
set rs=db.execute("delete from tblNewsletterLog where iMailingID=" & iId)
set rs=nothing
set rs=db.execute("delete from tblNewsletterMailing where iCustomerID=" & cId & " and iId=" & iId)
set rs=nothing
end if
end function
public property get nmbrReceivers(t)
if convertGetal(iId)<>0 then
dim rs
select case t
case 0
set rs=db.execute("select count(*) from tblNewsletterLog where iMailingID=" & iId)
case 1
set rs=db.execute("select count(*) from tblNewsletterLog where bRead=" & getSQLBoolean(true) & " and iMailingID=" & iId)
case 2
set rs=db.execute("select count(*) from tblNewsletterLog where bRead=" & getSQLBoolean(false) & " and iMailingID=" & iId)
end select
nmbrReceivers=rs(0)
set rs=nothing
else
nmbrReceivers=0
end if
end property
public sub showRead(t)
response.write "<table><tr><td align=left><ul>"
dim rs,sql
sql="select tblNewsletterCategorySubscriber.sName, tblNewsletterCategorySubscriber.sEmail from tblNewsletterCategorySubscriber where tblNewsletterCategorySubscriber.iId in (select tblNewsletterLog.iSubscriberID from tblNewsletterLog where "
select case t
case "0"
'do nothing
case "1"
sql=sql& " tblNewsletterLog.bRead=" & getSQLBoolean(true) & " and "
case "2"
sql=sql& " tblNewsletterLog.bRead=" & getSQLBoolean(false) & " and "
end select
sql=sql & " tblNewsletterLog.iMailingID=" & iId &")"
set rs=db.execute(sql)
while not rs.eof
response.write "<li>" & quotrep(rs("sName")) & " (<a href='mailto:" & quotrep(rs("sEmail")) & "'>" & quotrep(rs("sEmail")) & "</a>)</li>"
rs.movenext
wend 
set rs=nothing
response.write "</ul></td></tr></table>"
end sub
end class%>
