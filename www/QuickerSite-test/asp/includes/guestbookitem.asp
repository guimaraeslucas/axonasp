
<%class cls_guestbookitem
Public iId
Public sValue
Public sMessageBy
Public sMessageByEmail
Public dCreatedTS
Public bApproved
Public iGuestBookID
Public sKey
Public ip
Public sReply
Private p_guestbook
Private Sub Class_Initialize
iId=null
set p_guestbook=nothing
end sub
Private Sub Class_Terminate
set p_guestbook=nothing
end sub
Public sub  approve()
if convertGetal(iId)<> 0 then
dim rs
set rs=db.execute("update  tblGuestbookItem set bApproved=" & getSQLBoolean(true) & " where iId="& convertGetal(iId))
set rs=nothing
end if
end sub
Public function pickByKey(key)
dim sql, rs, id
sql = "select iId from tblGuestbookItem where sKey='" & left(cleanup(key),32) & "'"
set rs = db.execute(sql)
if not rs.eof then
id=rs("iId")
set RS = nothing
pick(id)
end if
end function
Public function guestbook
if p_guestbook is nothing then
set p_guestbook=new cls_guestbook
p_guestbook.pick(iGuestBookID)
end if
set guestbook=p_guestbook
end function
Public Function Pick(id)
dim sql, RS, sValue
if isNumeriek(id) then
sql = "select * from tblGuestbookItem where iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
sValue	= rs("sValue")
sMessageBy	= rs("sMessageBy")
sMessageByEmail	= rs("sMessageByEmail")
dCreatedTS	= rs("dCreatedTS")
bApproved	= rs("bApproved")
iGuestBookID	= rs("iGuestBookID")
sKey	= rs("sKey")
ip	= rs("ip")
sReply	= rs("sReply")
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
if isLeeg(sValue) then
check=false
message.AddError("err_mandatory")
exit function
end if
if isLeeg(sMessageBy) then
check=false
message.AddError("err_mandatory")
exit function
end if
if instr(lcase(sValue),"href=")<>0 then
check=false
message.AddError("err_mandatory")
exit function
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
'check block IP
if instr(guestbook.sBlockIP,UserIP) then
exit function
end if
set rs = db.GetDynamicRS
if isLeeg(iId) then
sKey=GeneratePassWord() & GeneratePassWord()
rs.Open "select * from tblGuestBookItem where 1=2"
rs.AddNew
rs("dCreatedTS")	= now()
rs("sKey")	= sKey
if guestbook.bRequireValidation then
bApproved	= false
else
bApproved	= true
end if
if not isLeeg(guestbook.sEmail) then
bascicLink=customer.sQSUrl & "/default.asp?redPage=" & encrypt(selectedPage.iId) & "&sKey="&sKey&"&pageAction=gbedit&ac="
dim mlinks,bascicLink
mlinks="<p>"
mlinks=mlinks&"<a href='" & bascicLink & "approve" & "'>Approve</a>   "
mlinks=mlinks&"<a href='" & bascicLink & "delete" & "'>Delete</a>   "
mlinks=mlinks&"<a href='" & bascicLink & "edit" & "'>Edit</a>"
mlinks=mlinks&"</p>"
dim wMail,wLink
set wMail=new cls_mail_message
wMail.receiver=guestbook.sEmail
wMail.subject="New entry for Guestbook " & guestbook.sName 
wMail.body=left(sanitize(sMessageBy),50) & "<br /><br />" & left(sanitize(sMessageByEmail),50) & "<br /><br/>" & left(sanitize(sValue),5000) & mlinks 
wMail.send
set wMail=nothing	 
end if
else
rs.Open "select * from tblGuestBookItem where iId="& iId
end if
rs("sValue")	= left(sValue,5000)
rs("sMessageBy")	= left(sMessageBy,50)
rs("sMessageByEmail")	= left(sMessageByEmail,50)
rs("iGuestBookID")	= iGuestBookID
rs("bApproved")	= bApproved
rs("ip")	= trim(ip)
rs("sReply")	= sReply
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
selectedPage.clearPageCache()
end function
public function remove
if convertGetal(iId)<>0 then
dim rs
set rs=db.execute("delete from tblGuestbookItem where iId="& convertGetal(iId))
set rs=nothing
end if
end function
end class%>
