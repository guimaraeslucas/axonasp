<!-- #include file="begin.asp"-->


<%bUseArtLoginTemplate=true
dim secondAdmin
set secondAdmin=customer.secondAdmin
dim sDefPW
if sha256(QS_defaultPW)=customer.adminPassword and not devVersion then sDefPW="admin"%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBackShort.asp"--><!-- #include file="bs_header.asp"--><%if Request.Form ("btnaction")="Login" then
If LCase(session("captcha")) <> LCase(Left(Request.Form("captcha"),4)) Then
' Add mandatory error if security image was not correct
message.AddError("err_captcha")
elseif logon.logon(sha256(Request.Form ("password"))) then
if isLeeg(request("bs_page")) then
Response.Redirect ("bs_default.asp")
else
if instr(request("bs_page"),"contactEdit.asp")<>0 then
session("bLoadCPinPP")="yes"
end if
Response.Redirect (request("bs_page"))
end if
else
message.AddError("err_login")
end if
application("bsLoginCount"&UserIP)=convertGetal(application("bsLoginCount"&UserIP))+1
end if
logon.lockBSAdmin()
dim nmbrAtt
nmbrAtt=convertGetal(application("bsLoginCount"&UserIP))+1%><p align=center><%=getSecurityWarning%></p><%if not isLeeg(sDefPW) then %><p align=center><b>The default password is "admin". You will be asked to change it in a second...</b></p><%end if%><form action="<%=QS_backsite_login_page%>" method="post" name="mainform"><input type="hidden" name="btnaction" value="Login" /><input type="hidden" name="bs_page" value="<%=server.HTMLEncode (request("bs_page"))%>" /><table align="center" cellpadding="2"><tr><td class="QSlabel"><%=l("password")%>:</td><td colspan="2"><input type="password" size="10" value="<%=sDefPW%>" name="password" maxlength="50" required /></td></tr><tr><td class="QSlabel"><%=l("captcha")%>:* </td><td><img src='includes/captcha.asp' alt="Captcha Image" /></td><td><input type="text" name="captcha" maxlength="4" size="6" autocomplete="off" required /></td></tr><tr><td class="QSlabel">&nbsp;</td><td colspan="2"><input class="art-button" type="submit"  name="dummy" value="<%=l("login")%>" /></td></tr></table></form><script type="text/javascript">document.mainform.password.focus();</script><p align=center><%=l("attempt")%> <font color=Red><b><%=nmbrAtt%><%if nmbrAtt=QS_number_of_allowed_attempts_to_login then Response.Write " - " & l("lastchance") & "!!!"%></b></font></p><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
