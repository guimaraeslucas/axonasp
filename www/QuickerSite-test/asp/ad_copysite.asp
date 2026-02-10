<!-- #include file="begin.asp"-->


<!-- #include file="ad_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%'site to be copied
dim siteToBeCopied
set siteToBeCopied=new cls_customer
siteToBeCopied.pick(convertGetal(Request("siteToBeCopiedID")))
siteToBeCopied.sName=""
siteToBeCopied.sURL=""
siteToBeCopied.dOnlineFrom=date()
siteToBeCopied.sAlternateDomains=""
customer.pick(siteToBeCopied.iId)
if convertBool(Request.Form ("postBack")) then
checkCSRF()
siteToBeCopied.sUrl	= convertStr(Request.Form ("sUrl"))
siteToBeCopied.sAlternateDomains	= convertStr(Request.Form ("sAlternateDomains"))
siteToBeCopied.sDescription	= convertStr(Request.Form ("sDescription"))
siteToBeCopied.siteName	= convertStr(Request.Form ("siteName"))
siteToBeCopied.siteTitle	= convertStr(Request.Form ("siteTitle"))
siteToBeCopied.copyRight	= convertStr(Request.Form ("copyRight"))
siteToBeCopied.keywords	= convertStr(Request.Form ("keywords"))
siteToBeCopied.googleAnalytics	= convertStr(Request.Form ("googleAnalytics"))
siteToBeCopied.language	= convertStr(Request.Form ("language"))
siteToBeCopied.sDatumFormat	= convertStr(Request.Form ("sDatumFormat"))
siteToBeCopied.webmaster	= convertStr(Request.Form ("webmaster"))
siteToBeCopied.webmasterEmail	= convertStr(Request.Form ("webmasterEmail"))
siteToBeCopied.sName	= convertStr(Request.Form ("siteName"))
siteToBeCopied.dOnlineFrom	= convertDateFromPicker(Request.Form ("dOnlineFrom"))
siteToBeCopied.bApplication	= convertBool(Request.Form ("bApplication"))
siteToBeCopied.bScanReferer	= convertBool(Request.Form ("bScanReferer"))
siteToBeCopied.bUserFriendlyURL	= convertBool(Request.Form ("bUserFriendlyURL"))
siteToBeCopied.bAllowStorageOutsideWWW	= convertBool(Request.Form ("bAllowStorageOutsideWWW"))
siteToBeCopied.bEnableMainRSS	= convertBool(Request.Form ("bEnableMainRSS"))
siteToBeCopied.defaultTemplate	= convertLng(decrypt(Request.Form ("defaultTemplate")))
if siteToBeCopied.check() then
siteToBeCopied.copy()
Response.Redirect ("ad_default.asp?newAccount=True&iId=" & siteToBeCopied.iId)
end if
end if
dim languageList
set languageList=new cls_languageListNew
dim dateFormatList
set dateFormatList=new cls_dateFormatList%><form action="ad_copysite.asp" name="mainform" method="post"><%=QS_secCodeHidden%><input type="hidden" name="siteToBeCopiedID" value="<%=siteToBeCopied.iID%>" /><input type="hidden" name="postback" value="<%=true%>" /><table align=center cellpadding="2"><tr><td class=header colspan=2>Copy site:</td></tr><tr><td class=QSlabel><%=l("nameoforganisation")%>:</td><td><input name="siteName" value="<%=quotRep(siteToBeCopied.siteName)%>" type=text size=40 maxlength=254 /></td></tr><tr><td class=QSlabel><%=l("url")%>:</td><td><input name="sUrl" value="<%=quotRep(siteToBeCopied.sUrl)%>" type=text size=40 maxlength=254 /></td></tr><%if siteToBeCopied.templates.count>0 then%><tr><td class="QSlabel"><%=l("defaulttemplate")%>:</td><td><select name="defaultTemplate"><option value=""><%=MYQS_name%></option><%=siteToBeCopied.showSelectedtemplate("option", siteToBeCopied.defaultTemplate)%></select></td></tr><%end if%><tr><td class=QSlabel><%=l("alternatedomains")%>:<br /><i><%=l("enterseplist")%></i></td><td><textarea name="sAlternateDomains" cols=40 rows=2><%=quotRep(siteToBeCopied.sAlternateDomains)%></textarea></td></tr><tr><td class=QSlabel><%=l("titleofsite")%>:</td><td><input name="siteTitle" value="<%=quotRep(siteToBeCopied.siteTitle)%>" type=text size=40 maxlength=254 /></td></tr><tr><td class=QSlabel><%=l("copyrightonsite")%>:<br /><i>(META-tag copyright)</i></td><td><input name="copyRight" value="<%=quotRep(siteToBeCopied.copyRight)%>" type=text size=40 maxlength=254 /></td></tr><tr><td class=QSlabel><%=l("keywords")%>:<br /><i>(META-tag keywords)</i></td><td><input name="keywords" value="<%=quotRep(siteToBeCopied.keywords)%>" type=text size=40 maxlength=2048 /></td></tr><tr><td class=QSlabel><%=l("description")%>:<br /><i>(META-tag description)</i></td><td><textarea name="sDescription" cols=40 rows=2><%=quotRep(siteToBeCopied.sDescription)%></textarea></td></tr><tr><td class=QSlabel><%=l("dateformat")%>:</td><td><select name="sDatumFormat"><%=dateFormatList.showSelected("option",siteToBeCopied.sDatumFormat)%></select></td></tr><tr><td class=QSlabel><%=l("languageofsite")%>:</td><td><select name="language"><%=languageList.showSelected("option",siteToBeCopied.language)%></select></td></tr><tr><td class=QSlabel><%=l("name")%> webmaster:<br /><i>(META-tag author)</i></td><td><input name="webmaster" value="<%=quotRep(siteToBeCopied.webmaster)%>" type=text size=40 maxlength=255 /></td></tr><tr><td class=QSlabel><%=l("email")%> webmaster:</td><td><input name="webmasterEmail" value="<%=quotRep(siteToBeCopied.webmasterEmail)%>" type=text size=40 maxlength=254 /></td></tr><tr><td class=QSlabel><%=l("since")%>:</td><td><input type="text" id="dOnlinefrom" name="dOnlinefrom" size="13" value="<%=convertEuroDate(siteToBeCopied.dOnlinefrom)%>" /><%=JQDatePicker("dOnlinefrom")%></td></tr><tr><td class=QSlabel>Main RSS?</td><td><input type=checkbox name=bEnableMainRSS STYLE="BORDER:0" value=1 <% if siteToBeCopied.bEnableMainRSS then Response.Write "checked"%> /></td></tr><tr><td class=QSlabel><%=l("allowapplications")%></td><td><input type=checkbox name=bApplication STYLE="BORDER:0" value=1 <% if siteToBeCopied.bApplication then Response.Write "checked"%> /></td></tr><tr><td class=QSlabel><%=l("allowuserfriendlyurls")%></td><td><input type=checkbox name=bUserFriendlyURL STYLE="BORDER:0" value=1 <% if siteToBeCopied.bUserFriendlyURL then Response.Write "checked"%> /></td></tr><tr><td class=QSlabel><%=l("allowstorageoutsidewww")%></td><td><input type=checkbox name=bAllowStorageOutsideWWW STYLE="BORDER:0" value=1 <% if siteToBeCopied.bAllowStorageOutsideWWW then Response.Write "checked"%> /></td></tr><tr><td class=QSlabel>Scan referrers</td><td><input type=checkbox name=bScanReferer STYLE="BORDER:0" value=1 <% if siteToBeCopied.bScanReferer then Response.Write "checked"%> /></td></tr><!--
<tr><td class=QSlabel>After creation of copy, execute url:</td><td><input type=text name=bExecute STYLE="BORDER:0" value=1 <% if siteToBeCopied.bScanReferer then Response.Write "checked"%> /></td></tr>--><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type="submit" value="<%=l("copy")%>" name="btnaction" /></td></tr></table></form><script type="text/javascript">document.mainform.siteName.focus()</script><%customer.pick(cId)%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
