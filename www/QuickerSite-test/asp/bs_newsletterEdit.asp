<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bNewsletter%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><!-- #include file="includes/urlenCodeJS.asp"--><%=getBOHeader(btn_Newsletter)%><%dim Newsletter
set Newsletter=new cls_Newsletter
dim postBack
postBack=convertBool(Request.Form("postBack"))
if request.form("btnaction")=l("delete") then
Newsletter.remove
Response.Redirect ("bs_NewsletterList.asp")
end if
dim previewName,previewEmail
if postBack then
Newsletter.getRequestValues()
checkCSRF()
if Newsletter.save then 
previewName=request.form("previewName")
previewEmail=request.form("previewEmail")
if request.form("btnaction")=l("send") then
Newsletter.send previewName,previewEmail,""
else
message.Add("fb_saveOK")
end if
end if
end if
if isLeeg(previewName) or isLeeg(previewEmail) then
previewName=customer.webmaster
previewEmail=customer.webmasterEmail
end if%><table align=center><tr><td align=center>-> <b><a href="bs_NewsletterList.asp"><%=l("back")%></a></b> <-</td></tr></table><form action="bs_NewsletterEdit.asp" method="post" name="mainform"><input type=hidden name=iNewsletterId value="<%=encrypt(Newsletter.iID)%>" /><input type=hidden name=postBack value="<%=true%>" /><%=QS_secCodeHidden%><table align=center cellpadding=3><tr><td colspan="2" class="header"><%=l("general")%>:</td></tr><tr><td class="QSlabel">Name:*</td><td colspan="2"><input type=text size=80 maxlength=150 name="sName" value="<%=quotRep(Newsletter.sName)%>" /></td></tr><%if isNumeriek(Newsletter.iID) then%><tr><td class=QSlabel><a name="MS"></a>Send preview to:</td><td><%if request.form("btnaction")=l("send") then%><span style="padding:5px;background-color:Yellow;color:Green"><strong>Preview was sent!</strong></span>  <%end if%><input type="text" name="previewName" value="<%=previewName%>" /> - <input type="text" name="previewEmail" value="<%=previewEmail%>" /><input class="art-button" type=submit name=btnaction value="<% =l("send")%>" /></td></tr><%end if%><tr><td class="QSlabel">From Email:*</td><td colspan="2"><input type=text size=80 maxlength=50 name="sFromEmail" value="<%=quotRep(Newsletter.sFromEmail)%>"></td></tr><tr><td class="QSlabel">From Name:*</td><td colspan="2"><input type=text size=80 maxlength=50 name="sFromName" value="<%=quotRep(Newsletter.sFromName)%>"></td></tr><tr><td class="QSlabel">Subject:*</td><td colspan="2"><input type=text size=80 maxlength=150 name="sSubject" value="<%=quotRep(Newsletter.sSubject)%>"></td></tr><tr><td class="QSlabel">Text:</td><td colspan="2"><%createFCKInstance Newsletter.sValue,"siteBuilder","sValue"%></td></tr><tr><td class="QSlabel">Body background-color:*</td><td colspan="2"><input type="text" id="sBodyBGColor" name="sBodyBGColor" value="<%=quotrep(Newsletter.sBodyBGColor)%>" /><%=JQColorPicker("sBodyBGColor")%></td></tr><tr><td class="QSlabel">Unsubscribe-link:*</td><td colspan="2"><input type=text size=45 maxlength=50 name="sUnsubscribeText" value="<%=quotRep(Newsletter.sUnsubscribeText)%>"> - this will be the text used in [NL_UNSUBSCRIBELINK]</td></tr><tr><td class="QSlabel">CSS for unsubscribe-link:</td><td colspan="2">style="<input type=text size=45 maxlength=250 name="sStyleLink" value="<%=quotRep(Newsletter.sStyleLink)%>" />"</td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name="btnaction" value="<% =l("save")%>" /> 
<%if isNumeriek(Newsletter.iID) then%><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>" /><%end if%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input type=checkbox value="1" name="bSaveAsTemplate" /> Save as default new newsletter</td></tr></table></form><table align=center><tr><td align=center>-> <b><a href="bs_NewsletterList.asp"><%=l("back")%></a></b> <-</td></tr></table><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
