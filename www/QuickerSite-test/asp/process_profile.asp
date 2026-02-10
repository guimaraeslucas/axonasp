
<%set ticket=new cls_ticket
ticket.pickByTicket(request("ac"))
if not logon.authenticatedIntranet and convertGetal(ticket.iId)=0 then
if isLeeg(request("ac")) then
Response.Redirect ("default.asp?iId="&encrypt(selectedPage.iId))
else
Response.Redirect ("default.asp?pageAction="& cloginIntranet&"&iId="&encrypt(getIntranetHomePage.iId)&"&strMessage=err_wrongactivationlink")
end if
end if
if convertGetal(ticket.iId)<>0 and isLeeg(Request.Form ("sEmail"))  then
logon.contact.sEmail=ticket.sEmail
end if
set contactFields=customer.contactFields(false)
select case Request.Form ("btnaction")
case l("save")
checkCSRF()
logon.contact.getRequestValues(contactFields)
'overrule double emails
if convertGetal(ticket.iId)<>0 and lcase(logon.contact.sEmail)=lcase(ticket.sEmail) then logon.contact.allowDE=true
if logon.contact.save(contactFields) then
message.Add("fb_saveOK")
response.cookies.item("sPw") = encrypt(logon.contact.sPw)
response.cookies.item("sEmail") = logon.contact.sEmail
if convertGetal(ticket.iId)<>0 then
ticket.remove()
if customer.bSendMailUponNewMember then
dim wMail, wLink
set wMail=new cls_mail_message
wMail.receiver=customer.sEmailNewRegistrations
wMail.subject=l("newprofilecreated") & " " & customer.surl
wLink=customer.sQSUrl & "/asp/"&QS_backsite_login_page&"?bs_page="& server.URLEncode ("bs_contactEdit.asp?iContactID="& encrypt(logon.contact.iId))
wMail.body="<p>Email: "& logon.contact.sEmail & "<br />Nickname: "& logon.contact.sNickName &"<br />Check: <a href=" & """" & wLink & """" & ">"&wLink&"</a><hr />" & getVisitorDetails & "</p>"
wMail.send
set wMail=nothing
end if
'default status
'to avoid double emails
logon.contact.pick(logon.contact.iId)
if not isLeeg(ticket.sEmail) and lcase(ticket.sEmail)<>lcase(logon.contact.sEmail) then logon.contact.sOrigEmail=ticket.sEmail
logon.contact.iStatus=convertGetal(customer.iDefaultStatus)
logon.contact.save(contactFields)
Response.Redirect ("default.asp?pageAction="& cWelcome &"&fbMessage="&server.urlencode("fb_saveOK"))
end if
end if
end select
set fields=logon.contact.fields
pageBody="<form name='contact' method='post' action='default.asp'>"
pageBody=pageBody&QS_secCodeHidden
pageBody=pageBody&"<input type='hidden' name='pageAction' value='"&cProfile&"' />"
pageBody=pageBody&"<input type='hidden' name='ac' value='"&ticket.sTicket&"' />"
pageBody=pageBody&customer.sExplProfile
pageBody=pageBody&"<table cellpadding=""2"" border=""0"" style=""border-style:none"">"
pageBody=pageBody&"<tr><td colspan='2'><b>"&l("logginin")&":</b></td></tr>"
pageBody=pageBody&"<tr>"
pageBody=pageBody&"<td class='QSlabel'>"& l("email") &":*</td>"
pageBody=pageBody&"<td><input required type='email' size='30' maxlength='50' name='sEmail' value=" & """" & sanitize(logon.contact.sEmail) & """" & " /></td>"
pageBody=pageBody&"</tr>"
pageBody=pageBody&"<tr>"
pageBody=pageBody&"<td class='QSlabel'>"& l("password") &":*</td>"
pageBody=pageBody&"<td><input required type=password size='30' maxlength='20' name='sPw' value=" & """" & sanitize(logon.contact.sPw) & """" & " /></td>"
pageBody=pageBody&"</tr>"
pageBody=pageBody&"<tr>"
pageBody=pageBody&"<td class='QSlabel'>"& l("nickname") &":*</td>"
pageBody=pageBody&"<td><input required type='text' size='30' maxlength='50' name='sNickName' value=" & """" & sanitize(logon.contact.sNickName) & """" & " /></td>"
pageBody=pageBody&"</tr>"
pageBody=pageBody&"<tr>"
pageBody=pageBody&"<td class='QSlabel'>"& l("allowprivatemessages") &"</td>"
pageBody=pageBody&"<td><input type='radio'  name='bGetEmailsFromSite' " & convertChecked(logon.contact.bGetEmailsFromSite) & " value='" & true & "' /> " & l("yes") &" <input " & convertChecked(not logon.contact.bGetEmailsFromSite) & " type='radio'  name='bGetEmailsFromSite' value='" & false & "' /> " & l("no") &"</td>"
pageBody=pageBody&"</tr>"
pageBody=pageBody&"<tr><td colspan='2'><hr /></td></tr>"
pageBody=pageBody&"<tr><td colspan='2'><b>"&l("contactdata")&":</b></td></tr>"
for each contactField in contactFields
if contactFields(contactField).bProfile then
pageBody=pageBody&"<tr>"
pageBody=pageBody&"<td class='QSlabel' width='160'>"
if contactFields(contactField).sType<>sb_comment then
pageBody=pageBody& contactFields(contactField).sFieldname
end if
if contactFields(contactField).bMandatory then
pageBody=pageBody&"*"
end if
pageBody=pageBody&"</td>"
pageBody=pageBody&"<td>"
select case contactFields(contactField).sType
case sb_text, sb_url, sb_email
pageBody=pageBody&"<input type='text' maxlength='300' size='30' name='" & encrypt(contactField) &"' value=" & """" & sanitize(fields(contactField)) & """" & " />"
case sb_richtext
pagebody=pageBody&dumpFCKInstance(fields(contactField),"siteBuilderRichText",encrypt(contactField))
case sb_textarea
pageBody=pageBody&"<textarea cols='29' rows='3' name='" & encrypt(contactField) &"'>" & quotRep(fields(contactField)) &"</textarea>"
case sb_checkbox
pageBody=pageBody&"<input type='checkbox' name='"& encrypt(contactField) &"' value='checked' "& fields(contactField) &" />"
case sb_select
pageBody=pageBody&"<select name='"& encrypt(contactField) &"'>"& contactFields(contactField).showSelected(fields(contactField)) &"</select>"
case sb_date
pageBody=pageBody&"<input type=""text"" id=""" & encrypt(contactField) & """ name=""" & encrypt(contactField) & """ value=""" & sanitize(convertDateToPicker(fields(contactField))) & """ />" & JQDatePicker(encrypt(contactField))
case sb_comment
pageBody=pageBody&contactFields(contactField).sValues
end select
pageBody=pageBody&"</td></tr>"
else
select case contactFields(contactField).sType
case sb_date
pageBody=pageBody&"<input type='hidden' name='"& encrypt(contactField) &"' value='"& convertDateToPicker(fields(contactField)) &"' />"
case else
pageBody=pageBody&"<input type='hidden' name='"& encrypt(contactField) &"' value="& """" & quotRep(fields(contactField)) & """" & " />"
end select
end if
next
pageBody=pageBody&"<tr><td class=QSlabel>&nbsp;</td><td>(*) "& l("mandatory") &"</td></tr>"
pageBody=pageBody&"<tr><td class=QSlabel>&nbsp;</td><td><input class=""art-button"" type=submit name=btnaction value="""&sanitize(l("save")) &""" /></td></tr>"
if customer.bUseAvatars and convertGetal(logon.contact.iId)<>0 then
pageBody=pageBody&"<tr><td class=QSlabel>&nbsp;</td><td><a href=""default.asp?pageAction=avataredit"" class=""QSPPAVATAR"">Upload Avatar</a></td></tr>"
end if
pageBody=pageBody&"</table>"
pageBody=pageBody&"</form>"
pageTitle=customer.intranetMyProfile%>
