

<%'Response.Write "<p align=center><br /><font color=Red><b>" & l("sendstarted") & "!</b></font><br /><font color=Red><b>"& l("notcloserefresh") &"!</b></font></p>"
'Response.Write "<p align=center><div id=counter align=center></div></p>"
'Response.Flush 
server.ScriptTimeout=10000
dim contactDict
set contactDict=server.CreateObject ("scripting.dictionary")
dim theMail, subject, body, mailobj
set mailobj=new cls_mail
set theMail=new cls_mail_message
theMail.receiverName=""
subject=hSubject
body=hBody
dim ccEmails,contactIDS
ccEmails=split(Request.Form ("ccEmails"),",")
contactIDS=split(Request.Form ("iContactIDM"),",")
'Response.Flush 
dim contactKey, contact
for contactKey=lbound(contactIDS) to ubound(contactIDS) 
set contact=new cls_contact
contact.pick(decrypt(contactIDS(contactKey)))
if not isLeeg(contact.sEmail) then
if not contactDict.Exists (contact.sEmail) then
On Error Resume Next
theMail.subject	= subject
theMail.body	= body
theMail.subject	= replace(theMail.subject,"["&l("email")&"]",contact.sEmail)
theMail.subject	= replace(theMail.subject,"[sPw]",contact.sPw)
theMail.body	= replace(theMail.body,"["&l("email")&"]",contact.sEmail)
theMail.body	= replace(theMail.body,"[sPw]",contact.sPw)
for each field in contactFields
select case contactFields(field).sType
case sb_date
theMail.subject=replace(theMail.subject,"["& contactFields(field).sFieldname &"]",convertStr(convertDateToPicker(contact.fields(field))))
theMail.body=replace(theMail.body,"["& contactFields(field).sFieldname &"]",convertStr(convertDateToPicker(contact.fields(field))))
case else
theMail.subject=replace(theMail.subject,"["& contactFields(field).sFieldname &"]",convertStr(contact.fields(field)))
theMail.body=replace(theMail.body,"["& contactFields(field).sFieldname &"]",convertStr(contact.fields(field)))
end select
next
theMail.receiver=contact.sEmail
if err.number=0 then
theMail.send()
contactDict.Add contact.sEmail,""
set contact=nothing
'verzend naar cc
dim ccRunner
for ccRunner=lbound(ccEmails) to ubound(ccEmails)
if CheckEmailSyntax(ccEmails(ccRunner)) then
if not isLeeg(ccEmails(ccRunner)) then
theMail.receiver=ccEmails(ccRunner)
theMail.send()
end if
end if
next
else
ErrorReport contact.sEmail & " - " & theMail.subject,err
end if
err.Clear()
On Error Goto 0
end if
end if
next
On Error Resume Next
mailobj.sBody=body
mailobj.sSubject=subject
mailobj.sBodyBGColor=hBodyBGColor
mailobj.save()
dim mailRS
set mailRS = db.GetDynamicRS
mailRS.Open "select * from tblMailContact where 1=2"
dim emailKey
for each emailKey in contactDict
mailRS.AddNew()
mailRS("iMailID")=mailobj.iId
mailRS("sEmail")=emailKey
mailRS.update()
next
Response.Redirect ("bs_massMailingFB.asp?counter="& contactDict.Count)
On Error Goto 0
 %>
