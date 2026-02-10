<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bForms%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Forms)%><%dim form
set form=new cls_form
dim postBack
postBack=convertBool(Request.Form("postBack"))
if postBack then
form.getRequestValues()
end if
if convertGetal(form.iCustomerID) <> convertGetal(cId) and convertGetal(form.iId)<>0 then
Response.End 
end if
select case Request.Form ("btnaction")
case l("save")
checkCSRF()
form.getRequestValues()
dim newForm
newForm=isLeeg(form.iId)
if form.save then 
if newForm then
Response.Redirect ("bs_formfields.asp?iFormID="& encrypt(form.iId))
else
Response.Redirect ("bs_formList.asp")
end if
end if
case l("delete")
checkCSRF()
form.remove
Response.Redirect ("bs_formList.asp")
end select
dim urlTypeShortList
set urlTypeShortList=new cls_urlTypeShortList%><!-- #include file="bs_backForm.asp"--><form action="bs_formEdit.asp" method="post" name="mainform"><%=QS_secCodeHidden%><input type="hidden" name="iFormID" value="<%=encrypt(form.iID)%>" /><input type="hidden" value="<%=true%>" name="postback" /><table align=center cellpadding="2"><tr><td colspan=2 class="header"><%=l("general")%>:</td></tr><tr><td class="QSlabel"><%=l("name")%>:*</td><td><input type=text size=40 maxlength=250 name="sName" value="<%=sanitize(form.sName)%>" /></td></tr><tr><td class=QSlabel><%=l("introform")%>:</td><td><textarea cols=40 rows=5 name="sIntro"><%=quotRep(form.sIntro)%></textarea></td></tr><tr><td class=QSlabel><%=l("buttonsend")%>:*</td><td><input type=text size=20 maxlength=255 name="sButton" value="<%=quotRep(form.sButton)%>" /></td></tr><tr><td class=QSlabel><%=l("buttonreset")%>:</td><td><input type=text size=20 maxlength=255 name="sReset" value="<%=quotRep(form.sReset)%>" />&nbsp;<i><span style="font-size:0.8em">(<%=l("leaveblankifnotneeded")%>)</span></i></td></tr><tr><td class=QSlabel><%=l("formqaalign")%>:*</td><td><input type=radio value="<%=QS_QleftAright%>" name="sQAalign" <%if form.sQAalign=QS_QleftAright or isLeeg(form.sQAalign) then Response.Write "checked"%> />&nbsp;<%=l("qleftaright")%><br /><input type=radio value="<%=QS_QtopAbottom%>" name="sQAalign" <%if form.sQAalign=QS_QtopAbottom then Response.Write "checked"%> />&nbsp;<%=l("QtopAbottom")%><br /><input type=radio value="<%=QS_QrightAleft%>" name="sQAalign" <%if form.sQAalign=QS_QrightAleft or isLeeg(form.sQAalign) then Response.Write "checked"%> />&nbsp;<%=l("qrightaleft")%></td></tr><tr><td class=QSlabel><%=l("setcookie")%></td><td><input type=checkbox name="bCookie" value="1" <%=convertChecked(form.bCookie)%> /></td></tr><tr><td class=QSlabel><%=l("verifyhumanity")%></td><td><input type=checkbox name="bCaptcha" value="1" <%=convertChecked(form.bCaptcha)%> />&nbsp;<i>(<a href="http://en.wikipedia.org/wiki/Captcha" target=captcha>CAPTCHA</a>)</i></td></tr>


<tr><td class=QSlabel><%=l("feedbackform")%>:</td>
<td><textarea cols=40 rows=5 name="sFeedback"><%=quotRep(form.sFeedback)%></textarea>
<br /><span>tip: use <input type="text" value="[QS_COPYSUBMISSION]" size="22" onclick="javascript:this.select();" /> to include a copy of the submission in the autoreply</span>

</td></tr>
<tr><td class=QSlabel><%=l("redirecttourl")%>:</td><td><select name="sRedirectPrefix"><%=urlTypeShortList.showSelected("option",form.sRedirectPrefix)%></select>&nbsp;<input type=text size=40 maxlength=255 name="sRedirect" value="<%=quotRep(form.sRedirect)%>" /></td></tr><tr><td class=QSlabel><%=l("useautoresponse")%></td><td><input type=checkbox onclick="javascript:document.mainform.submit();" name="bAutoResponder" value="1" <%=convertChecked(form.bAutoResponder)%> /></td></tr><%if form.bAutoResponder then%><tr><td colspan=2 class=header><%=l("settingsautoresponse")%>:</td></tr><tr><td class=QSlabel><%=l("subject")%>:*</td><td><input type=text size=40 maxlength=255 name="sAutoResponseSubject" value="<%=quotRep(form.sAutoResponseSubject)%>" /></td></tr><tr><td class=QSlabel><%=l("fromname")%>:*</td><td><input type=text size=40 maxlength=255 name="sAutoResponseFromName" value="<%=quotRep(form.sAutoResponseFromName)%>" /></td></tr><tr><td class=QSlabel><%=l("fromemail")%>:*</td><td><input type=text size=40 maxlength=255 name="sAutoResponseFromEmail" value="<%=quotRep(form.sAutoResponseFromEmail)%>" /></td></tr><tr><td class=QSlabel><%=l("autoresponse")%>:*</td><td><textarea cols=40 rows=5 name="sAutoResponse"><%=quotRep(form.sAutoResponse)%></textarea>

<br /><span>tip: use <input type="text" value="[QS_COPYSUBMISSION]" size="22" onclick="javascript:this.select();" /> to include a copy of the submission in the autoreply</span>

</td></tr><%end if%><tr><td class=QSlabel><%=l("sendemail")%></td><td><input type=checkbox onclick="javascript:document.mainform.submit();" name="bSendEmail" value="1" <%=convertChecked(form.bSendEmail)%> /></td></tr><%if form.bSendEmail then%><tr><td colspan=2 class=header><%=l("emailsettings")%>:</td></tr><tr><td class=QSlabel><%=l("subject")%>:*</td><td><input type=text size=40 maxlength=255 name="sSubject" value="<%=quotRep(form.sSubject)%>" /></td></tr><tr><td class=QSlabel><%=l("autoresponse")%>:</td><td><textarea cols=40 rows=5 name="sAutoResponseWebmaster"><%=quotRep(form.sAutoResponseWebmaster)%></textarea></td></tr><tr><td class=QSlabel><%=l("receiver")%>:*<br /><i><%=l("enterseplist")%></i></td><td><textarea cols=40 rows=5 name="sTo"><%=quotRep(form.sTo)%></textarea></td></tr><tr><td class=QSlabel><%=l("includeattachments")%></td><td><input type=checkbox name="bAttachFiles" value="1" <%=convertChecked(form.bAttachFiles)%> /></td></tr><%end if%><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=btnaction value="<% =l("save")%>" /><%if isNumeriek(form.iID) then%><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>" /><br /><br /><%end if%></td></tr><%if customer.bApplication then 'idee nog niet rijp%><tr><td class=QSlabel valign=top><%=l("scriptuponsubmission")%></td><td><textarea cols=70 rows=20 name="sScriptUponSubmission"><%=quotRep(form.sScriptUponSubmission)%></textarea></td></tr><%end if%></table></form><!-- #include file="bs_backForm.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
