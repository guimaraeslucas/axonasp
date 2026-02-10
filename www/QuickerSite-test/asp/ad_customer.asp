<!-- #include file="begin.asp"-->


<!-- #include file="ad_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%dim customerObj
set customerObj=new cls_customer
customerObj.pick(request("iId"))
dim customerList
set customerList=new cls_customerList
dim customers
customers=customerList.count
if Request.Form ("btnaction")=l("save") then
checkCSRF()
dim newAccount
newAccount=false
if isLeeg(customerObj.iId) then
'one time setup
customerObj.adminPassword	= sha256(QS_defaultPW)
customerObj.webmasterEmail	= "your@email.com"
newAccount=true
elseif lcase(convertStr(customerObj.sAlternateDomains))<>lcase(Request.Form ("sAlternateDomains")) or  lcase(convertStr(customerObj.sURL))<>lcase(Request.Form ("sURL")) then
customerObj.saveBindings=true
customerObj.sOrigURL=customerObj.sURL
end if
customerObj.sName	= convertStr(Request.Form ("sName"))
customerObj.sURL	= convertStr(Request.Form ("sURL"))
customerObj.dOnlineFrom	= convertDateFromPicker(Request.Form ("dOnlineFrom"))
customerObj.bApplication	= convertBool(Request.Form ("bApplication"))
customerObj.bScanReferer	= convertBool(Request.Form ("bScanReferer"))
customerObj.sAlternateDomains	= convertStr(Request.Form ("sAlternateDomains"))
customerObj.bUserFriendlyURL	= convertBool(Request.Form ("bUserFriendlyURL"))
customerObj.bAllowStorageOutsideWWW = convertBool(Request.Form ("bAllowStorageOutsideWWW"))
customerObj.bEnableMainRSS	= convertBool(Request.Form ("bEnableMainRSS"))
customerObj.bMonitor	= convertBool(Request.Form ("bMonitor"))
customerObj.bEnableNewsletters	= convertBool(Request.Form ("bEnableNewsletters"))
customerObj.bUseCachingForPages	= convertBool(request.form("bUseCachingForPages"))
customerObj.bListItemPic	= convertBool(request.form("bListItemPic"))
customerObj.bShoppingCart	= convertBool(request.form("bShoppingCart"))

customerObj.SMTPSERVER	= convertStr(Request.Form ("C_SMTPSERVER"))
customerObj.SMTPPORT	= convertStr(Request.Form ("C_SMTPPORT"))
customerObj.SMTPUSERNAME	= convertStr(Request.Form ("C_SMTPUSERNAME"))
customerObj.SMTPUSERPW	= convertStr(Request.Form ("C_SMTPUSERPW"))
customerObj.SENDUSING	= convertStr(Request.Form ("C_SENDUSING"))
customerObj.SMTPUSESSL = convertStr(Request.Form ("C_SMTPUSESSL")) 

if customerObj.save() then
if not isLeeg(trim(Request.Form ("adminPassword"))) then
customerObj.resetDBConn=false
if customerObj.saveAdminPW(Request.Form ("adminPassword")) then
customerObj.resetDBConn=false

Response.Redirect ("ad_default.asp?newAccount="&newAccount&"&iId="&customerObj.iId)
end if
else
Response.Redirect ("ad_default.asp?newAccount="&newAccount&"&iId="&customerObj.iId)
end if
end if
end if
if Request.Form ("btnaction")=l("delete") then
checkCSRF()
customerObj.remove()
Response.Redirect ("ad_default.asp")
end if%><form action="ad_customer.asp" name="mainform" method="post"><%=QS_secCodeHidden%><input type=hidden name=iId value="<%=customerObj.iID%>" /><table align=center cellpadding="2"><tr><td class=header colspan=2><%=l("general")%>:</td></tr><tr><td class=QSlabel><%=l("nameoforganisation")%>:*</td><td><input type=text value="<%=quotRep(customerObj.sName)%>" name=sName size=40 maxlength=250 /></td></tr><tr><td class=QSlabel><%=l("url")%>:*</td><td><input type=text value="<%=quotRep(customerObj.sURL)%>" name=sURL size=40 maxlength=250 /><br />-> <i><%=l("includevd")%></i></td></tr><tr><td class=QSlabel><%=l("since")%>:</td><td><input type="text" id="dOnlinefrom" name="dOnlinefrom" value="<%=convertEuroDate(customerObj.dOnlinefrom)%>" /><%=JQDatePicker("dOnlinefrom")%></td></tr><tr><td class=QSlabel>Main RSS?</td><td><input type=checkbox name=bEnableMainRSS STYLE="BORDER:0" value=1 <% if customerObj.bEnableMainRSS then Response.Write "checked"%> /></td></tr><tr><td class=QSlabel><%=l("allowapplications")%></td><td><input type=checkbox name=bApplication STYLE="BORDER:0" value=1 <% if customerObj.bApplication then Response.Write "checked"%> /></td></tr>

<tr><td class=QSlabel><%=l("allowuserfriendlyurls")%></td><td><input type=checkbox name=bUserFriendlyURL STYLE="BORDER:0" value=1 <% if customerObj.bUserFriendlyURL then Response.Write "checked"%> /></td></tr>
<tr><td class=QSlabel>Use Picture uploads for list items?</td><td><input type=checkbox name=bListItemPic STYLE="BORDER:0" value=1 <% if customerObj.bListItemPic then Response.Write "checked"%> /></td></tr>
<tr><td class=QSlabel>Use Shopping cart?</td><td><input type=checkbox name=bShoppingCart STYLE="BORDER:0" value=1 <% if customerObj.bShoppingCart then Response.Write "checked"%> /></td></tr>
<tr><td class=QSlabel><%=l("allowstorageoutsidewww")%></td><td><input type=checkbox name=bAllowStorageOutsideWWW STYLE="BORDER:0" value=1 <% if customerObj.bAllowStorageOutsideWWW then Response.Write "checked"%> /></td></tr><tr><td class=QSlabel>Scan referrers</td><td><input type=checkbox name=bScanReferer STYLE="BORDER:0" value=1 <% if customerObj.bScanReferer then Response.Write "checked"%> /></td></tr>
<tr><td class=QSlabel>Monitor?</td><td><input type=checkbox name=bMonitor STYLE="BORDER:0" value=1 <% if customerObj.bMonitor then Response.Write "checked"%> /></td></tr><tr><td class=QSlabel>Enable newsletter?</td><td><input type=checkbox name=bEnableNewsletters STYLE="BORDER:0" value=1 <% if customerObj.bEnableNewsletters then Response.Write "checked"%> /></td></tr><tr><td class=QSlabel>Use caching for pages?</td><td><input type=checkbox name=bUseCachingForPages STYLE="BORDER:0" value=1 <% if customerObj.bUseCachingForPages then Response.Write "checked"%> /></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type="submit" value="<%=l("save")%>" name="btnaction" /><%if isNumeriek(customerObj.iId)  then


if customerObj.iId<>cId then%>&nbsp;<input class="art-button" onclick="javascript: return confirm('<%=l("areyousure")%>');" type="submit" value="<%=l("delete")%>" name="btnaction" /><%end if%><%end if%></td></tr>


<%if isNumeriek(customerObj.iId) then%><tr><td class=QSlabel>&nbsp;</td><td>&nbsp;</td></tr><tr><td class=QSlabel><%=l("totalpassword")%></td><td><%=jaNeen(not isLeeg(customerObj.sTotalPW))%></td></tr><tr><td class=QSlabel><%=l("createdon")%>:</td><td><%=convertEuroDate(customerObj.dCreatedTS)%></td></tr><tr><td class=QSlabel><%=l("updatedon")%>:</td><td><%=convertEuroDate(customerObj.dUpdatedTS)%></td></tr><tr><td class=QSlabel><%=l("resetstatisticson")%>:</td><td><%=convertEuroDate(customerObj.dResetStats)%></td></tr><tr><td class=QSlabel>Reset Backsite Password:</td><td><input type=text value="" name="adminPassword" size=20 maxlength=30 /><br /><i>(Leave empty in case no reset is needed)</i></td></tr><tr><td class=QSlabel>&nbsp;</td><td><a target="resetAPP<%=customerObj.iId%>" href="<%=customerObj.sQSUrl%>/asp/ad_resetApp.asp?apw=<%=sha256(C_ADMINPASSWORD)%>">Reset Application</a></td></tr><%end if%><tr><td class=QSlabel>alternative smtp-server:</td><td><fieldset><legend>smtp server settings</legend><table><tr><td class=QSlabel>smtp server:</td><td><input size="40" maxlength="50" type=text name=C_SMTPSERVER value="<%=customerObj.SMTPSERVER%>" /></td></tr><tr><td class=QSlabel>server port:</td><td><input size="7" maxlength="10" type=text name=C_SMTPPORT value="<%=customerObj.SMTPPORT%>" /> (default: 25)</td></tr><tr><td class=QSlabel>smtp user:</td><td><input size="40" maxlength="50" type=text name=C_SMTPUSERNAME value="<%=customerObj.SMTPUSERNAME%>" /></td></tr><tr><td class=QSlabel>smtp password:</td><td><input size="40" maxlength="150" type=text name=C_SMTPUSERPW value="<%=customerObj.SMTPUSERPW%>" /></td></tr>

<tr><td class=QSlabel valign="top">send using</td><td><input size="2" maxlength="1" type=text name=C_SENDUSING value="<%=customerObj.SENDUSING%>" />1: local server | 2: external server | 3: exchange server</td></tr>

<tr><td class=QSlabel valign="top">require SSL?</td><td><input size="2" maxlength="1" type=text name=C_SMTPUSESSL value="<%=customerObj.SMTPUSESSL%>" />0: no | 1: yes</td></tr>

</table></fieldset></td></tr><tr><td class=QSlabel valign="top"><%=l("alternatedomains")%>:<br /><i><%=l("enterseplist")%></i></td><td><textarea name="sAlternateDomains" cols=80 rows=10><%=quotRep(customerObj.sAlternateDomains)%></textarea></td></tr></table></form><%if customers>0 then%><!-- #include file="ad_back.asp"--><%end if%><%if convertGetal(customerObj.iId)=0  then%><script type="text/javascript">document.mainform.sName.focus();</script><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
