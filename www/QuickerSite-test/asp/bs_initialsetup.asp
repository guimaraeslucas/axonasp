<!-- #include file="begin.asp"-->


<%bUseArtLoginTemplate=true
'nodig voor redirection in bs_security.asp
blockDefaultPW=false%><!-- #include file="bs_security.asp"--><%if sha256(QS_defaultPW)<>customer.adminPassword then Response.redirect("bs_admin.asp")
sub getRequestForm()
customer.sUrl	= Request.Form ("sUrl")
customer.siteName	= Request.Form ("siteName")
customer.siteTitle	= Request.Form ("siteTitle")
customer.copyRight	= Request.Form ("copyRight")
customer.language	= Request.Form ("language")
customer.webmaster	= Request.Form ("webmaster")
customer.webmasterEmail	= Request.Form ("webmasterEmail")
customer.sDatumFormat	= Request.Form ("sDatumFormat")
end sub
customer.resetDBConn=false
if Request.Form ("btnaction")="saveSetupAdmin" then
checkCSRF()
getRequestForm()
if isLeeg(Request.form("start")) then
message.AddError("err_mandatory")
end if
if Request.Form("adminPassword")<>Request.Form("adminPasswordBis") then
message.AddError("pwnomatch")
end if
if customer.save() and not message.hasErrors then
if customer.saveAdminPW(Request.Form("adminPassword")) then
if Request.Form ("start")="clean" then
customer.bRemoveTemplatesOnSetup=false
customer.reset()
getRequestForm()
customer.save()
end if
removeApplication()
Response.Redirect ("bs_default.asp")
end if
end if
else
dim hhhtttppp
If LCase(Request.ServerVariables("HTTPS")) = "on" Then
hhhtttppp="https://"
else
hhhtttppp="http://"
end if
customer.sUrl=hhhtttppp & Request.ServerVariables ("http_host") & C_VIRT_DIR
end if
dim languageList
set languageList=new cls_languageListNew
dim dateFormatList
set dateFormatList=new cls_dateFormatList%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><p align=center><%=l("thankyouforusingqs")%><br /><%=l("initialsetup")%></p><form action="bs_initialsetup.asp" name="mainform" method="post" onsubmit="javascript:return confirm('<%=l("areyousure")%>');"><input type=hidden name="btnaction" value="saveSetupAdmin" /><%=QS_secCodeHidden%><table align=center><tr><td class=QSlabel><%=l("nameoforganisation")%>:*</td><td><input name="siteName" value="<%=quotRep(customer.siteName)%>" type=text size=30 maxlength=254 /></td></tr><tr><td class=QSlabel><%=l("url")%>:*</td><td><input name="sUrl" value="<%=quotRep(customer.sUrl)%>" type=text size=30 maxlength=254 /></td></tr><tr><td class=QSlabel><%=l("titleofsite")%>:</td><td><input name="siteTitle" value="<%=quotRep(customer.siteTitle)%>" type=text size=30 maxlength=254 /></td></tr><tr><td class=QSlabel><%=l("copyrightonsite")%>:<br /><i>(META-tag copyright)</i></td><td><input name="copyRight" value="<%=quotRep(customer.copyRight)%>" type=text size=30 maxlength=254 /></td></tr><tr><td class=QSlabel><%=l("dateformat")%>:</td><td><select name="sDatumFormat"><%=dateFormatList.showSelected("option",customer.sDatumFormat)%></select></td></tr><tr><td class=QSlabel><%=l("languageofsite")%>:</td><td><select name="language"><%=languageList.showSelected("option",customer.language)%></select></td></tr><tr><td class=QSlabel><%=l("name")%> webmaster:<br /><i>(META-tag author)</i></td><td><input name="webmaster" value="<%=quotRep(customer.webmaster)%>" type=text size=30 maxlength=254 /></td></tr><tr><td class=QSlabel><%=l("email")%> webmaster:</td><td><input name="webmasterEmail" value="<%=quotRep(customer.webmasterEmail)%>" type=text size=30 maxlength=254 /></td></tr><tr><td class=QSlabel><%=l("password")%> backsite:*</td><td><input name="adminPassword" value="<%=sanitize(Request.Form ("adminPassword"))%>" type=password size=8 maxlength=30 /></td></tr><tr><td class=QSlabel><%=l("retypepassword")%>:*</td><td><input name="adminPasswordBis" value="<%=sanitize(Request.Form ("adminPasswordBis"))%>" type=password size=8 maxlength=30 /></td></tr><tr><td class=QSlabel valign=top><%=l("pleaseselect")%>:*</td><td><input type=radio name=start value="sample" <%if Request.Form("start")="sample" then Response.Write "checked"%> /><%=l("startfromss")%><br /><input type=radio name=start value="clean" <%if Request.Form("start")="clean" then Response.Write "checked"%> /><%=l("startfromes")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=dummy value="<%=l("save")%>" /></td></tr></table></form><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
