<!-- #include file="begin.asp"-->


<%dim contact
set contact=new cls_contact
if not logon.authenticatedIntranet then Response.End 
if not contact.bGetEmailsFromSite then Response.End 
if convertBool(Request.Form ("postback")) then
If LCase(session("CAPTCHA")) <> LCase(Left(Request.Form("CAPTCHA"),4)) Then
message.AddError("err_captcha")
elseif isLeeg(Request.Form ("body")) then
message.AddError("err_mandatory")
else
dim myMail
set myMail=new cls_mail_message
myMail.receiver=contact.sEmail
myMail.receiverName=contact.sNickName
myMail.subject="Private message from " & logon.contact.sNickName & " (" & customer.surl & ")"
myMail.body=linkURLS(Request.Form ("body"))
myMail.fromemail=logon.contact.sEmail
myMail.fromname=logon.contact.sNickName
myMail.send
'send CC?
if convertBool(Request.Form ("cc")) then
myMail.receiver=logon.contact.sEmail
myMail.receiverName=logon.contact.sNickName
myMail.send
end if
set myMail=nothing
Response.Redirect ("fs_mailcontact.asp?mailsent=true&iContactID=" & encrypt(contact.iId))
end if
end if%><!-- #include file="includes/commonHeader.asp"--><body style="margin:7px;color:#000;background-color:#FFF;background-image:none" onload="javascript:window.focus();"><%if Request.QueryString ("mailsent")="true" then%><p align="center"><%=l("fb_mailsent")%></p><%else%><p align="center"><%=replace(l("sendprivatemessage"),"[NICKNAME]","<b>"& contact.sNickName &"</b>",1,-1,1)%></p><form action="fs_mailcontact.asp" method="post"><input type="hidden" name="postback" value="<%=true%>" /><input type="hidden" name="iContactID" value="<%=encrypt(contact.iId)%>" /><table  style="background-color:#FFF" cellspacing="3" align="center" width="500"><tr><td style="text-align:right"><i><%=l("from")%></i></td><td><%=logon.contact.sNickName%>&nbsp;<i>(<%=logon.contact.sEmail%>)</i></td></tr><tr><td style="text-align:right"><i><%=l("to")%></i></td><td><%=contact.sNickName%></td></tr><tr><td valign="top" style="text-align:right"><i><%=l("yourmessage")%></i>*</td><td><textarea name="body" cols=60 rows=7><%=sanitize(Request.Form ("body"))%></textarea></td></tr><tr><td valign="top" style="text-align:right">&nbsp;</td><td><input type="checkbox" name="cc" value="<%=true%>" checked />&nbsp;<i><%=l("sendacopyto")%>&nbsp;<%=logon.contact.sEmail%></i></td></tr><tr><td style="text-align:right"><i><%=l("captcha")%>:</i>* </td><td><img style="vertical-align: middle;" alt='' src='includes/captcha.asp' />&nbsp;<input type="text" size="6" name="captcha" maxlength="4" autocomplete="off" required /></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input type="submit" value="<%=l("send")%>" /></td></tr></table></form><%end if%><%=message.showAlert()%><%=cPopup.dumpJS()%> 
</body></html><%cleanupASP%>
