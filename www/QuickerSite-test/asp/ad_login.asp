<!-- #include file="begin.asp"-->


<%bUseArtLoginTemplate=true
dim secondAdmin
set secondAdmin=customer.secondAdmin
if request("btnaction")= "loginAdmin" then
If LCase(session("captcha")) <> LCase(Left(Request.Form("captcha"),4)) Then
' Add mandatory error if security image was not correct
message.AddError("err_captcha")
elseif logon.logonAdmin (Encrypt(Request.Form ("password"))) then
Response.Redirect ("ad_default.asp")
else
message.AddError("err_login")
end if
application("adminLoginCount"&UserIP)=convertGetal(application("adminLoginCount"&UserIP))+1
end if
logon.lockAdmin()
dim nmbrAtt
nmbrAtt=convertGetal(application("adminLoginCount"&UserIP))+1%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><p align=center><b>ADMIN Login</b></p><p align=center><%=getSecurityWarning%></p><form action="<%=QS_admin_login_page%>" method="post" name="mainform"><input type="hidden" name="btnaction" value="loginAdmin" /><table align="center"><tr><td class="QSlabel"><%=l("password")%>:</td><td colspan="2"><input type="password" size="10" name="password" required /></td></tr><tr><td class="QSlabel"><%=l("captcha")%>:* </td><td><img src="includes/captcha.asp" alt="Captcha Image" /></td><td><input type="text" name="captcha" maxlength="4" size="6" autocomplete="off" required /></td></tr><tr><td class="QSlabel">&nbsp;</td><td><input type="submit"  class="art-button"  name="dummy" value="<%=l("login")%>" /></td></tr></table></form><script type="text/javascript">document.mainform.password.focus();</script><p align=center><%=l("attempt")%> <font color=Red><b><%=nmbrAtt%><%if nmbrAtt=QS_number_of_allowed_attempts_to_login then Response.Write " - "& l("lastchance") &"!!!"%></b></font></p><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
