
<%class cls_newsletterCategory
Public iId,sName,sSignupForm,sUnsubscribeFB,sUnsubscribeFBTitle,sWelcomeMessage,bRequireBoth,sErrorMessage,sNotifEmail
Private Sub Class_Initialize
On Error Resume Next
sSignupForm="<form action=""default.asp"" method=""post"" id=""nl{ID}"" name=""nl{ID}"">" & vbcrlf & "{NL_HIDDENFIELD_DO_NOT_TOUCH}" & vbcrlf & "<table border=""0"">" & vbcrlf & "<tr>" & vbcrlf & "<td>Name:</td>" & vbcrlf & "<td><input name=""{N}"" type=""text"" value=""{NL_NAME}"" /></td>" & vbcrlf & "</tr>" & vbcrlf & "<tr>" & vbcrlf & "<td>Email:</td>" & vbcrlf & "<td><input name=""{E}"" type=""email"" value=""{NL_EMAIL}"" required /></td>" & vbcrlf & "</tr>" & vbcrlf & "<tr>" & vbcrlf & "<td> </td>" & vbcrlf & "<td><input class=""art-button"" onclick=""{AJAX}"" type=""button"" value=""Submit"" name=""dummy"" /></td>" & vbcrlf & "</tr>" & vbcrlf & "</table>" & vbcrlf & "</form>"
sUnsubscribeFB="<p>Dear [NL_NAME],</p><p>We have removed your email address <strong>[NL_EMAIL]</strong> from our mailing list.</p><p>" & customer.webmaster &"</p>"
sUnsubscribeFBTitle="So sorry to see you go..."
sWelcomeMessage="<p>Hello <b>[NL_NAME]</b>! Your email address <b>[NL_EMAIL]</b> is now added to our mailing list.</p>"
bRequireBoth=true
sErrorMessage="<font color=""Red"">Enter both name and (valid) email please...</font>"
sNotifEmail=customer.webmasterEmail
pick(decrypt(replace(request("iNewsletterCategoryID"),"NLC","",1,-1,1)))
On Error Goto 0
end sub
Public Function getRequestValues()
sName	= left(request.form("sName"),50)
sSignupForm	= request.form("sSignupForm")
sUnsubscribeFB	= request.form("sUnsubscribeFB")
sUnsubscribeFBTitle	= left(request.form("sUnsubscribeFBTitle"),50)
sWelcomeMessage	= request.form("sWelcomeMessage")
bRequireBoth	= convertBool(request.form("bRequireBoth"))
sErrorMessage	= left(request.form("sErrorMessage"),255)
sNotifEmail	= left(request.form("sNotifEmail"),50)
end Function
Public Function Pick(id)
dim sql, rs
if isNumeriek(id) then
sql = "select * from tblNewsletterCategory where iCustomerID="&cid&" and iId=" & id
set rs = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
sName	= rs("sName")
sSignupForm	= rs("sSignupForm")
sUnsubscribeFB	= rs("sUnsubscribeFB")
sUnsubscribeFBTitle	= rs("sUnsubscribeFBTitle")
sWelcomeMessage	= rs("sWelcomeMessage")
bRequireBoth	= rs("bRequireBoth")
sErrorMessage	= rs("sErrorMessage")
sNotifEmail	= rs("sNotifEmail")
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
if isLeeg(sUnsubscribeFB) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(sUnsubscribeFBTitle) then
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
rs("sSignupForm") 	= sSignupForm
rs("iCustomerID")	= cId
rs("sUnsubscribeFB")	= sUnsubscribeFB
rs("sUnsubscribeFBTitle")	= sUnsubscribeFBTitle
rs("sWelcomeMessage")	= sWelcomeMessage
rs("bRequireBoth")	= bRequireBoth
rs("sErrorMessage")	= sErrorMessage
rs("sNotifEmail")	= sNotifEmail
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
set rs=db.execute("select count(*) from tblNewsletterCategorySubscriber where bActive=" & getSQLBoolean(true) & " and iCategoryID=" & iId)
nmbrSubscribers=convertgetal(rs(0))
set rs=nothing
end property
public function build()
dim cQSSECNL
if isLeeg(session("QSSECNL" & iId)) then
cQSSECNL=generatePassword & generatePassword & generatePassword
session("QSSECNL" & iId)=cQSSECNL
else
cQSSECNL=session("QSSECNL" & iId)
end if
build="<div id=""divNLC" & iId & """>"
'build=build & "<p>name: " & request.querystring("sName") & "</p>"
build=build&sSignupForm
build=replace(build,"{ID}",iId,1,-1,1)
build=replace(build,"{NL_HIDDENFIELD_DO_NOT_TOUCH}","<input type=""hidden"" name=""QSSEC"" value=""" & cQSSECNL & """ />",1,-1,1)
build=replace(build,"{N}","n" & iId,1,-1,1)
build=replace(build,"{E}","e" & iId,1,-1,1)
build=replace(build,"{NL_NAME}",left(quotrep(trim(request.querystring("sName"))),50),1,-1,1)
build=replace(build,"{NL_EMAIL}",left(quotrep(trim(request.querystring("sEmail"))),50),1,-1,1)
build=replace(build,"{AJAX}","javascript:qs_div='NLC" & iId & "';mode='fb';getSub(nl"&iId&"." & "n" & iId &".value,nl"&iId&"." & "e" & iId &".value,'" & encrypt(iId) & "',nl" & iId & ".QSSEC.value);return false;",1,-1,1)
build=build&"</div>"
response.write err.description
end function
Public sub registerNameAndEmail
if convertStr(session("QSSECNL" & iId))<>request.querystring("cQSSECNL") or isLeeg(session("QSSECNL" & iId)) then
exit sub
end if
dim bContinue
bContinue=true
'checking
if bRequireBoth then
if isLeeg(request.querystring("sName")) then
bContinue=false
end if
end if
if not checkEmailSyntax(lcase(trim(request.querystring("sEmail")))) then
bContinue=false
end if
if bContinue then
dim rs
set rs=db.execute("select * from tblNewsletterCategorySubscriber where sEmail='" & trim(lcase(request.querystring("sEmail"))) &"' and iCategoryID=" & iId)
if rs.eof then
set rs=nothing
set rs=db.getDynamicRS
rs.open "select * from tblNewsletterCategorySubscriber where 1=2"
rs.AddNew()
rs("iCustomerID")=cId
rs("sEmail")=left(lcase(trim(request.querystring("sEmail"))),50)
rs("sName")=left(trim(request.querystring("sName")),50)
rs("bActive")=true
rs("sKey")=lcase(generatePassword & generatePassword & generatePassword)
rs("iCategoryID")=iId
rs("dAdded")=date()
rs.update()
rs.close
set rs=nothing
if not isLeeg(sNotifEmail) then
dim ncEMail
set ncEMail=new cls_mail_message
ncEMail.receiver=sNotifEmail
ncEMail.subject="New signup for email list '" & sName & "': " & quotrep(request.querystring("sEmail"))
ncEMail.body="email: " & quotrep(request.querystring("sEmail")) & "<br />" & "name: " & quotrep(request.querystring("sName"))
ncEMail.send
set ncEMail=nothing
end if
else
db.execute("update tblNewsletterCategorySubscriber set bActive=" & getSQLBoolean(true) & " where sEmail='" & trim(lcase(request.querystring("sEmail"))) &"' and iCategoryID=" & iId)
if bRequireBoth then
if not isLeeg(trim(request.querystring("sName"))) then
db.execute("update tblNewsletterCategorySubscriber set sName='" & left(replace(trim(request.querystring("sName")),"'","''",1,-1,1),50) & "' where sEmail='" & trim(lcase(request.querystring("sEmail"))) &"' and iCategoryID=" & iId)
end if
end if
end if
sWelcomeMessage=replace(sWelcomeMessage,"[NL_NAME]",quotrep(trim(request.querystring("sName"))),1,-1,1)
sWelcomeMessage=replace(sWelcomeMessage,"[NL_EMAIL]",quotrep(trim(request.querystring("sEmail"))),1,-1,1)
response.write sWelcomeMessage
session("QSSECNL" & iId)=generatePassword & generatePassword & generatePassword
else
response.write "<div id=""nc""" & iId & ">" & sErrorMessage & "</div>" & build
end if
end sub
end class%>
