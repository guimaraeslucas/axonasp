<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%if isleeg(customer.sMOBUrl) then
customer.sMOBUrl="http://"
end if
if request.form("postback")="1" then
session("checkMobileBrowser")=""
customer.sMOBBrowsers	= trim(request.form("sMOBBrowsers"))
customer.sMOBUrl	= trim(request.form("sMOBUrl"))
customer.iDefaultMobileTemplate	= convertGetal(decrypt(request.form("iDefaultMobileTemplate")))
end if
if request.form("btnaction")=l("save") then
customer.save()
response.redirect("bs_mobileSetup.asp?fbMessage=fb_saveOK")
end if%><p align="center">This site can be configured to use a specific template, or redirect a user to a specific url in case a specific browser is used.<br />This might be useful in case you want to offer a mobile version of your website to smartphones.</p><p align="center">Use this feature with care. Do not use it unless you know very well how and why you need it.</p><form method="post" action="bs_mobileSetup.asp" name="popupForm"><input type=hidden name=postback value="1" /><table align="center" cellpadding="2"><tr><td class=QSlabel valign=top>Enter separate keywords<br />to match the HTTP_USER_AGENT:</td><td><textarea style="width:400px" cols=40 rows=10 name="sMOBBrowsers"><%=sanitize(customer.sMOBBrowsers)%></textarea></td></tr><tr><td class=QSlabel>Use template:</td><td><select name=iDefaultMobileTemplate><option value=0>Please select</option><%=customer.showSelectedtemplate("option", customer.iDefaultMobileTemplate)%></select></td></tr><tr><td class=QSlabel>Redirection url:<br />(in case you don't use a template!)</td><td><input style="width:400px" type=text size=50 maxlength=255 name="sMOBUrl" value="<%=sanitize(customer.sMOBUrl)%>" /></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit value="<%=l("save")%>" name="btnaction" /></td></tr><tr><td  class=QSlabel valign=top><br /><br />Most recent HTTP_USER_AGENTS:</td><td style="width:500px"><br /><br /><%=replace(application("QSmostrecentUA"),vbcrlf,"<br /><br />",1,-1,1)%></td></tr></table></form><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
