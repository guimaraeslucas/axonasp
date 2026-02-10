
<%class cls_ticket
Public iId
Public sEmail
Public sTicket
Public sVisitorDetails
Public dCreatedTS
Private Sub Class_Initialize
On Error Resume Next
pick(decrypt(request("iTicketID")))
On Error Goto 0
end sub
Public function pickByTicket(ac)
dim sql, RS
sql = "select iId from tblContactRegistration where iCustomerID="&cid&" and sTicket='" & left(cleanup(ac),255) & "'"
set RS = db.execute(sql)
if not rs.eof then pick(rs("iId"))
set RS = nothing
end function
Public Function Pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblContactRegistration where iCustomerID="&cid&" and iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
sEmail	= rs("sEmail")
sTicket	= rs("sTicket")
dCreatedTS	= rs("dCreatedTS")
sVisitorDetails	= rs("sVisitorDetails")
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
if isLeeg(sEmail) then
check=false
message.AddError("err_email")
elseif not CheckEmailSyntax(sEmail) then
message.AddError("err_email")
check=false
else
sEmail=lcase(sEmail)
end if
dim SessionCAPTCHA
SessionCAPTCHA = Trim(Session("CAPTCHA"))
Session("CAPTCHA") = vbNullString
if Len(SessionCAPTCHA) < 1 then
message.AddError("err_captcha")
check=false
end if
if lcase(convertStr(SessionCAPTCHA)) <> lcase(convertStr(request.Form("captcha"))) then
message.AddError("err_captcha")
check=false
end if
if Check then
'check double email
dim rs
set rs=db.execute("select iId from tblContactRegistration where iCustomerID="& cId &" and sEmail='" & sEmail & "'")
if not rs.eof then
message.AddError("err_activationlink")
check=false
end if
set rs=db.execute("select iId from tblContact where iStatus>"&cs_silent&" and iCustomerID="& cId &" and sEmail='" & sEmail & "'")
if not rs.eof then
message.AddError("err_doubleemail")
check=false
end if
end if
End Function
Public Function SaveAndSend()
if check() then
SaveAndSend=true
else
SaveAndSend=false
exit function
end if
dim rs
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblContactRegistration where 1=2"
rs.AddNew
rs("dCreatedTS")=now()
else
rs.Open "select * from tblContactRegistration where iId="& iId
end if
sTicket	= GeneratePassWord()&GeneratePassWord()
rs("sTicket")	= sTicket
rs("sEmail")	= sEmail
rs("sVisitorDetails")	= getVisitorDetails()
rs("iCustomerID")	= cId
rs.Update 
iId = convertGetal(rs("iId"))
'send
sendTicket()
rs.close
Set rs = nothing
end function
public function remove
if not isLeeg(iId) then
dim rs
set rs=db.execute("delete from tblContactRegistration where iCustomerID="&cId&" and iId="& iId)
set rs=nothing
end if
end function
public sub sendTicket
dim theMail
set theMail=new cls_mail_message
theMail.receiver=sEmail
theMail.subject=treatConstants(customer.sSubjectMailTicket,true)
theMail.body=replace(convertStr(treatConstants(customer.sMailTicket,true)),"[QS_Intranet:ActivationLink]",activationlink,1,-1,1)
theMail.send
set theMail=nothing
end sub
private property get activationlink
activationlink=customer.sQSUrl & "/default.asp?pageAction="&cProfile&"&ac="&sticket
end property
end class%>
