<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Setup)%><%=getBOSetupMenu(btn_Security)%><%dim secAdmin
set secAdmin=new cls_secondAdmin
select case Request.Form ("btnaction")
case cSaveAdminPW
checkCSRF()
secAdmin.getPasswordValues()
if not isleeg(Request.form("sPasswordConfirm")) then
if secAdmin.save() then message.Add ("fb_saveOK")
else
message.AddError("err_mandatory")
end if
end select%><form action="bs_applyTotalPWSA.asp" method="post" name="mainform3"><input type="hidden" name="btnaction" value="<%=cSaveAdminPW%>" /><%=QS_secCodeHidden%><table align=center width=500 cellpadding="2"><tr><td class=QSlabel style="width:250px"><%=l("password")%> backsite:</td><td><input name="sPassword" value="" type=password size=15 maxlength=30 /></td></tr><tr><td class=QSlabel style="width:250px"><%=l("confirmpassword")%> backsite:</td><td><input name="sPasswordConfirm" value="" type=password size=15 maxlength=30 /></td></tr><tr><td class=QSlabel style="width:250px">&nbsp;</td><td><input class="art-button" type=submit onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("save")%>" name=dummy  /></td></tr></table></form><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
