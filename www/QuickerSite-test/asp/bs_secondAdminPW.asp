<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Setup)%><%=getBOSetupMenu(btn_Security)%><%if logon.currentPW=customer.secondAdmin.sPassword then Response.Redirect ("bs_default.asp")
dim secAdmin
set secAdmin=new cls_secondAdmin
select case Request.Form ("btnaction")
case l("save")
checkCSRF()
secAdmin.getPasswordValues()
if secAdmin.save() then  Response.Redirect ("bs_secondAdmin.asp")
end select%><form action="bs_secondAdminPW.asp" method="post" name="mainform"><%=QS_secCodeHidden%><p align=center><%=l("explsecondadmin")%></p><p align=center><b><%=l("step12ndadmin")%></b></p><input type="hidden" name="postback" value="<%=true%>" /><table align=center cellpadding="2"><tr><td class="QSlabel"><%=l("secondadminpw")%>:*</td><td><input name="sPassword" value="" type="password" size="15" maxlength="30" /></td></tr><tr><td class="QSlabel"><%=l("confirmpassword")%>:*</td><td><input name="sPasswordConfirm" value="" type="password" size="15" maxlength="30" /></td></tr><tr><td class="QSlabel">&nbsp;</td><td><input class="art-button" type="submit" onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("save")%>" name="btnaction" /></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr></table></form><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
