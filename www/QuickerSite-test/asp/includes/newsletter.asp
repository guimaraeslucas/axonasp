
<%class cls_newsletter
Public iId,sName,sValue,iCustomerID,bOnline,sSubject,sFromEmail,sFromName,sUnsubscribeText,sUnsubscribeFB,sUnsubscribeFBTitle,sBodyBGColor, sStyleLink
Private Sub Class_Initialize
On Error Resume Next
bConstantsLoaded=false
sValue="<table width=""580"" cellspacing=""0"" cellpadding=""20"" align=""center"" style=""border:6px solid #FFFFFF;font-family: Verdana; font-size: 8pt;""><tbody><tr><td height=""150"" bgcolor=""#375C71"" style=""border-bottom:6px solid #FFFFFF;font-family: Verdana; font-size: 12pt; text-align: right;color:#D3E1E9""><em><strong>Newsletter</strong></em></td></tr> <tr><td bgcolor=""#efefef"" style=""border-bottom:6px solid #FFFFFF;font-family: Verdana; font-size: 8pt;"">  <p>Dear [NL_NAME],</p>  <p>Your text here...</p>  <p>Best regards,<br />" & customer.webmaster & "</p>  </td>  </tr>  <tr>  <td bgcolor=""#B8C3CF"" style=""font-family: Verdana; font-size: 7pt; text-align: center;"">[NL_UNSUBSCRIBELINK]</td>  </tr>  </tbody></table>"
sSubject="Subject line"
sFromName=customer.webmaster
sFromEmail=customer.webmasterEmail
sUnsubscribeText="Unsubscribe from this newsletter"
sBodyBGColor="#DEE3E8"
sStyleLink="color:#FFFFFF"
bOnline=true
pick(decrypt(request("iNewsletterID")))
if isLeeg(iId) then
if not isLeeg(customer.sNLTemplate) then
on error resume next
dim nltemp
nltemp=split(customer.sNLTemplate,"|||###qsdelimiter###|||")
sSubject=nltemp(0)
sValue=nltemp(1)
sFromEmail=nltemp(2)
sFromName=nltemp(3)
sUnsubscribeText=nltemp(4)
sUnsubscribeFB=nltemp(5)
sUnsubscribeFBTitle=nltemp(6)
sBodyBGColor=nltemp(7)
sStyleLink=nltemp(8)
sName=nltemp(9)
on error goto 0
end if
end if
On Error Goto 0
end sub
Public property get sGetUrl
if not isLeeg(sUrl) and sUrl<>"http://" then
sGetUrl=sUrl
else
sGetUrl=customer.sQSUrl & "/default.asp?pageAction=showPopup&iPopupID=" & encrypt(iId)
end if
end property
Public Function getRequestValues()
sName	= left(trim(Request.Form ("sName")),150)
bOnline	= convertBool(request.form("bOnline"))
sValue	= removeEmptyP(request.form("sValue"))
sSubject	= left(trim(Request.Form ("sSubject")),150)
sFromName	= left(trim(Request.Form ("sFromName")),50)
sFromEmail	= left(trim(Request.Form ("sFromEmail")),50)
sUnsubscribeText	= left(trim(Request.Form ("sUnsubscribeText")),50)
sBodyBGColor	= left(trim(Request.Form ("sBodyBGColor")),50)
sStyleLink	= left(trim(request.form("sStyleLink")),255)
if convertBool(request.form("bSaveAsTemplate")) then
customer.sNLTemplate=sSubject & "|||###qsdelimiter###|||" & sValue & "|||###qsdelimiter###|||" & sFromEmail & "|||###qsdelimiter###|||" & sFromName & "|||###qsdelimiter###|||" & sUnsubscribeText & "|||###qsdelimiter###|||" & sUnsubscribeFB & "|||###qsdelimiter###|||" & sUnsubscribeFBTitle & "|||###qsdelimiter###|||" & sBodyBGColor & "|||###qsdelimiter###|||" & sStyleLink & "|||###qsdelimiter###|||" & sName
customer.save()
end if
end Function
Public Function Pick(id)
dim sql, rs
if isNumeriek(id) then
sql = "select * from tblNewsletter where iCustomerID="&cid&" and iId=" & id
set rs = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
sName	= rs("sName")
bOnline	= rs("bOnline")
sValue	= rs("sValue")
sSubject	= rs("sSubject")
sFromEmail	= rs("sFromEmail")
sFromname	= rs("sFromname")
sUnsubscribeText	= rs("sUnsubscribeText")
sBodyBGColor	= rs("sBodyBGColor")
sStyleLink	= rs("sStyleLink")
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
if isLeeg(sUnsubscribeText) then
check=false
message.AddError("err_mandatory")
end if
if not checkEmailSyntax(sFromEmail) then
check=false
message.AddError("err_email")
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
rs.Open "select * from tblNewsletter where 1=2"
rs.AddNew
else
rs.Open "select * from tblNewsletter where iId="& iId
end if
rs("sName")	= sName
rs("bOnline")	= bOnline
rs("iCustomerID")	= cId
rs("sValue")	= sValue
rs("sSubject")	= sSubject
rs("sFromname")	= sFromname
rs("sFromEmail")	= sFromEmail
rs("sUnsubscribeText")	= sUnsubscribeText
rs("sBodyBGColor")	= sBodyBGColor
rs("sStyleLink")	= sStyleLink
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
end function
public function remove
if not isLeeg(iId) then
dim rs
set rs=db.execute("delete from tblNewsletter where iId=" & iId)
set rs=nothing
end if
end function
public function copy()
if isNumeriek(iId) then
iId=null
save()
end if
end function
private bConstantsLoaded,sCopyVT
Public function send(n,e,k)
n=convertStr(n)
e=convertStr(e)
k=convertStr(k)
dim copyV,copyS
copyV=sValue
copyS=sSubject
sOWBodyBGColor=sBodyBGColor
if not bConstantsLoaded then
sCopyVT=treatConstants(copyV,true)
bConstantsLoaded=true
copyV=sCopyVT
else
copyV=sCopyVT
end if
bLoadConstants=false
dim NLEMail
set NLEMail=new cls_mail_message
NLEmail.fromemail=sFromEmail
NLEmail.fromname=sFromName
NLEMail.receiver=e
NLEMail.receiverName=n
copyS=replace(copyS,"[NL_NAME]",n,1,-1,1)
copyS=replace(copyS,"[NL_EMAIL]",e,1,-1,1)
NLEMail.subject=copyS
copyV=replace(copyV,"[NL_NAME]",n,1,-1,1)
copyV=replace(copyV,"[NL_EMAIL]",e,1,-1,1)
copyV=replace(copyV,"[NL_UNSUBSCRIBELINK]","<a style=""" & sStyleLink & """ href=""" & customer.sQSUrl & "/default.asp?pageAction=unsubscribe&e=" & k & """>" & sUnsubscribeText & "</a>",1,-1,1)
NLEMail.body=copyV
NLEMail.send
set NLEMail=nothing
end function
end class%>
