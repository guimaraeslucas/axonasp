<!-- #include file="asp/begin.asp"-->
<%printReplies=true
pagetoemail=true%><!-- #include file="asp/process.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" dir="<%=tdir%>" lang="en-US" xml:lang="en">
<head>
<title><%=selectedPage.showTitle%></title>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=<%=QS_CHARSET%>" />
<link rel="stylesheet" TYPE="text/css" href="<%=C_DIRECTORY_QUICKERSITE%>/css/qs_<%=tdir%>.css" title="QSStyle" />
<script type="text/javascript" src="<%=C_DIRECTORY_QUICKERSITE%>/js/sorttable.js"></script>
<%if not isLeeg(selectedPage.sApplication) then Response.End%>
</head>

<%if not isLeeg(selectedPage.sApplication) then Response.End
dim mailSent, trimEmail
mailSent=false
if Request.Form ("btnAction")="sendPage" then
trimEmail=trim(Request.Form("sEmail"))
checkCSRF()
session("QS_CMS_email")=trimEmail
' Check for security image
If LCase(session("CAPTCHA")) <> LCase(Left(Request.Form("CAPTCHA"),4)) or Request.Form("CAPTCHA")="" Then
' Add mandatory error if security image was not correct
message.AddError("err_captcha")
elseif not CheckEmailSyntax(trimEmail) then
'geen email ingevuld!
message.AddError("err_email")
else
dim mailBody, mailHeader, mailFooter
if isleeg(pagetoemailbody) then
'header
mailHeader="<html><head><title>"& pageTitle &"</title>"& css() &"</head>"
mailHeader=mailHeader&"<body class=main><table style='FONT-FAMILY:Arial;FONT-SIZE:10pt;' align=left width=555><tr><td style='FONT-FAMILY:Arial;FONT-SIZE:10pt;'>"
'body
mailBody="<b>" & pageTitle & "</b><hr />" & pagebody
'footer
mailFooter="</td></tr></table></body></html>"
mailbody=mailHeader&mailBody&mailFooter
else
mailbody=pagetoemailbody
end if
dim pageEMail
set pageEMail=new cls_mail_message
pageEMail.receiver=trimEmail
pageEMail.subject=selectedPage.replaceBlocks(pageTitle & " " & l("on") & " " & customer.sUrl)
pageEMail.body=selectedPage.replaceBlocks(mailbody)
pageEMail.send
set pageEMail=nothing
mailSent=true
end if
end if%><%if mailSent then%><body class="main" onLoad="javascript: window.focus();document.mailFB.btnClose.focus();"><form action="default.asp" method="post" name="mailFB"><table align="center" cellpadding="8"><tr><td align="center"><%=l("pageIsSent")%></td></tr><tr><td align="center"><input type="button" class="button" onClick="javascript: window.close();" name="btnClose" value="<%=l("close")%>"></td></tr></table></form><%else%><body class="main" onload="javascript: window.focus();document.mainform.sEmail.focus();"><form action="mailPage.asp" method="post" name="mainform"><%=QS_secCodeHidden%><input type="hidden" name="btnAction" value="sendPage"><input type="hidden" value="<%=encrypt(selectedPage.iId)%>" name="iId"><table align="center" cellpadding="8"><tr><td colspan="3"><%=l("thePage")%> '<b><%=pageTitle%></b> <%=l("willBeSent")%></td></tr><tr><td class="label"><%=l("email")%>:*</td><td colspan="2"><input required size="30" type="email" name="sEmail" value="<%=quotRep(session("QS_CMS_email"))%>"></td></tr><tr><td class="label"><%=l("captcha")%>:* </td><td><img src='<%=C_DIRECTORY_QUICKERSITE%>/asp/includes/captcha.asp' alt='CAPTCHA' /></td>        <td><input type="text" size="6" name="captcha" maxlength="4" required></td></tr><tr><td class="label">&nbsp;</td><td colspan="2">(*) <%=l("mandatory")%></td></tr><tr><td class="label">&nbsp;</td><td colspan="2"><input type="submit" name="dummy" value="<%=l("send")%>" class="button"></td></tr></table></form><%=message.showAlert()%><%end if%></body></html><%cleanUPASP%>
