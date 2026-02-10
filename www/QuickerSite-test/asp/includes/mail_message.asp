
<%class cls_mail_message
public receiver
public receiverName
public subject
public body
public attachments
public fromemail
public fromname
private sub class_initialize
receiverName=customer.siteName
set attachments=server.CreateObject ("scripting.dictionary")
end sub
Public function send
On Error Resume next
subject	= treatConstants(convertTo_iso_8859_1(subject),true)
body	= wrapInHTML(prepareforEmail(treatConstants(body,true)),subject)
body=replace(body,">" & vbcrlf &".",">.",1,-1,1)
if not isLeeg(sAddImageUrl)	then
body=replace(body,"</body>","<img width=""1"" height=""1"" src=""" & sAddImageUrl & """ /></body>",1,-1,1)
end if

if isLeeg(fromemail) then fromemail=customer.webmasterEmail
if isLeeg(fromname) then fromname=customer.siteName

fromname=RemoveHTML(fromname)

if not isLeeg(fromname) then
	fromname=replace(fromname,"""","'",1,-1,1)
	fromname=replace(fromname,","," ",1,-1,1)
	fromname=trim(fromname)
end if

if receiverName=customer.siteName then receiverName=receiver
receiverName=RemoveHTML(receiverName)

dim d_SMTPSERVER,d_SMTPPORT,d_SMTPUSERNAME,d_SMTPUSERPW,d_SENDUSING,d_SMTPUSESSL
d_SMTPSERVER	= C_SMTPSERVER
d_SMTPPORT	= C_SMTPPORT
d_SMTPUSERNAME	= C_SMTPUSERNAME
d_SMTPUSERPW	= C_SMTPUSERPW
d_SENDUSING	= C_SENDUSING
d_SMTPUSESSL = convertBool(QSCDO_smtpusessl)
if not isLeeg(customer.SMTPSERVER) then
d_SMTPSERVER	= customer.SMTPSERVER
d_SMTPPORT	= customer.SMTPPORT
d_SMTPUSERNAME	= customer.SMTPUSERNAME
d_SMTPUSERPW	= customer.SMTPUSERPW
d_SENDUSING	= customer.SENDUSING
d_SMTPUSESSL = convertBool(customer.SMTPUSESSL)
end if
dim myMail, fileKey
select case lcase(convertStr(C_MAILCOMPONENT))
case "persits.mailsender"
set myMail=Server.CreateObject("Persits.MailSender")
myMail.host=d_SMTPSERVER
myMail.Port=d_SMTPPORT
myMail.from=fromemail


myMail.fromName=fromname
myMail.addAddress receiver,receiverName
myMail.subject=subject
myMail.IsHTML=true
if d_SMTPUSERNAME<>"" then
myMail.username=d_SMTPUSERNAME
myMail.password=d_SMTPUSERPW
end if
'attachments
for each fileKey in attachments
myMail.AddAttachment attachments(fileKey)
next 
myMail.body=body
myMail.send
set myMail=nothing
case "cdo.message"
set myMail=Server.CreateObject("CDO.Message")
'configuration object
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")=d_SMTPSERVER
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=d_SMTPPORT 
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout")=60
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")=d_SENDUSING
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = "C:\Inetpub\mailroot\Pickup"

if d_SMTPUSERNAME<>"" then
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")=1
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername")=d_SMTPUSERNAME
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword")=d_SMTPUSERPW
end if
if d_SMTPUSESSL then
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
else
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
end if

myMail.Configuration.Fields.Update
myMail.Bodypart.ContentMediaType = "text/html"
myMail.Bodypart.Charset = QS_CHARSET
myMail.Bodypart.ContentTransferEncoding = "quoted-printable"
myMail.Subject=subject
myMail.To=receiverName&"<"&receiver&">"
myMail.From=fromname&"<"& fromemail &">"
myMail.ReplyTo=fromemail
'attachments
for each fileKey in attachments
myMail.AddAttachment attachments(fileKey)
next 
myMail.HtmlBody=body
myMail.HTMLBodypart.Charset = QS_CHARSET
myMail.send
'response.write server.htmlencode(myMail.HtmlBody)
'response.end 
set MyMail=nothing
case "cdonts.newmail"
set myMail=server.CreateObject("CDONTS.NewMail")
myMail.From= fromname & "<" & fromemail & ">"
myMail.To=receiverName&"<"&receiver&">"
myMail.Subject=subject
myMail.BodyFormat=0
myMail.MailFormat=0
'attachments
for each fileKey in attachments
myMail.AttachFile attachments(fileKey)
next 
myMail.Body=body
myMail.Send
set myMail=nothing
case "jmail.message"
set myMail=Server.CreateObject("jmail.message")
myMail.silent = true
myMail.AddRecipient receiver, receiverName
myMail.From = fromemail
myMail.FromName = fromname
myMail.Subject = subject

dim bodyNOhtml
bodyNOhtml=body

myMail.body=removehtml(bodyNOhtml)

myMail.HTMLBody=body
'attachments
for each fileKey in attachments
myMail.AddCustomAttachment fileKey, attachments(fileKey), false
next 
if d_SMTPUSERNAME<>"" then
myMail.MailServerUserName=d_SMTPUSERNAME
myMail.MailServerPassWord=d_SMTPUSERPW
end if
myMail.send(d_SMTPSERVER)
set MyMail=nothing
case "smtpsvg.mailer"
set myMail = Server.CreateObject("SMTPsvg.Mailer")
myMail.RemoteHost = d_SMTPSERVER ' Specify a valid SMTP server
myMail.FromAddress = fromemail ' Specify sender's address (only one)
myMail.FromName = fromname ' Optionally specify sender's name
myMail.AddRecipient  receiverName, receiver  'Specify recipient
myMail.CustomCharSet = "utf-8" 
for each fileKey in attachments
myMail.AddAttachment  = attachments(fileKey)    'How to add an attachment
next 
myMail.BodyText = body
myMail.ContentType   = "text/html"
'--Finally, send it
if myMail.SendMail then
'do nothing...
end if
set myMail=nothing
case else
Response.Write "Mail component '<b>" & lcase(convertStr(C_MAILCOMPONENT)) & "</b>' is NOT SUPPORTED! Check asp/config/web_config.asp!"
Response.End 
end select
On Error Goto 0
end function
end class%>
