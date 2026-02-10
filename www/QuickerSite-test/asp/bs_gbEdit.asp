<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bGuestbook%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><!-- #include file="includes/urlenCodeJS.asp"--><%=getBOHeader(btn_gb)%><%dim guestbook
set guestbook=new cls_guestbook
dim postBack
postBack=convertBool(Request.Form("postBack"))
if postBack then
guestbook.getRequestValues()
end if
select case Request.Form ("btnaction")
case l("save")
checkCSRF()
guestbook.getRequestValues()
if guestbook.save then 
Response.Redirect ("bs_gbList.asp")
end if
case l("delete")
checkCSRF()
guestbook.remove
Response.Redirect ("bs_gbList.asp")
end select%><!-- #include file="bs_gbBack.asp"--><form action="bs_gbEdit.asp" method="post" name="mainform"><input type=hidden name=iGBId value="<%=encrypt(guestbook.iID)%>" /><input type="hidden" value="<%=true%>" name=postback /><%=QS_secCodeHidden%><table align=center cellpadding="2"><tr><td colspan="2" class="header"><%=l("general")%>:</td></tr><tr><td class="QSlabel">Name:*</td><td colspan="2"><input type=text size=50 maxlength=50 name="sName" value="<%=quotRep(guestbook.sName)%>" /></td></tr><tr><td class=QSlabel>Open <%=l("from")%>:</td><td><input type="text" id="dOnlinefrom" name="dOnlinefrom" value="<%=convertEuroDate(guestbook.dOnlineFrom)%>" /><%=JQDatePicker("dOnlinefrom")%><i><%=l("untill")%>:</i> 
<input type="text" id="dOnlineUntil" name="dOnlineUntil" value="<%=convertEuroDate(guestbook.dOnlineUntil)%>" /><%=JQDatePicker("dOnlineUntil")%></td></tr><tr><td class="QSlabel"><%=l("code")%>:*</td><td>[QS_GUESTBOOK:<input type="text" size="10" maxlength="45" name="sCode" value="<%=quotRep(guestbook.sCode)%>" />]</td></tr><tr><td class="QSlabel">Full template:*</td><td><textarea rows="12" name="sFullTemplate" cols="90"><%=quotRep(guestbook.sFullTemplate)%></textarea></td></tr><tr><td class="QSlabel">Template form:*</td><td><textarea rows="12" name="sTemplateForm" cols="90"><%=quotRep(guestbook.sTemplateForm)%></textarea></td></tr><tr><td class="QSlabel">Template error:*</td><td><textarea rows="12" name="sTemplateErr" cols="90"><%=quotRep(guestbook.sTemplateErr)%></textarea></td></tr><tr><td class="QSlabel">Template Item:*</td><td><textarea rows="12" name="sTemplate" cols="90"><%=quotRep(guestbook.sTemplate)%></textarea></td></tr><tr><td class="QSlabel">Template Reply:*</td><td><textarea rows="12" name="sTemplateReply" cols="90"><%=quotRep(guestbook.sTemplateReply)%></textarea></td></tr><tr><td class="QSlabel">Show entries:</td><td><input type=radio name="sSortby" value="recentfirst" <%if guestbook.sSortby="recentfirst" then Response.Write " checked='checked' "%> /> Recent first 
<input type=radio name="sSortby" value="" <%if guestbook.sSortby="" then Response.Write " checked='checked' "%> /> Oldest first 
</td></tr><tr><td class="QSlabel">Show entries by:*</td><td><select name="iPaging"><%=numberList(5,500,5,guestbook.iPaging)%></select></td></tr><tr><td class="QSlabel">Entries need validation?</td><td><input type="radio" value="<%=true%>" name="bRequireValidation" <%if guestbook.bRequireValidation then Response.Write "checked='checked'"%> /> <%=l("yes")%> <input type="radio" value="<%=false%>" name="bRequireValidation" <%if not guestbook.bRequireValidation then Response.Write "checked='checked'"%> /> <%=l("no")%></td></tr><tr><td class="QSlabel">Note for items to be approved:</td><td><input type=text size=90 maxlength=255 name="sWarningApproval" value="<%=quotRep(guestbook.sWarningApproval)%>" /></td></tr><tr><td class="QSlabel">Send notification email to:</td><td><input type=email size=50 maxlength=50 name="sEmail" value="<%=quotRep(guestbook.sEmail)%>" /></td></tr><tr><td class="QSlabel" valign="top">Blocked IP:<br /><i>(enter-separated list)</i></td><td><textarea rows="8" name="sBlockIP" cols="20"><%=quotRep(guestbook.sBlockIP)%></textarea></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name="btnaction" value="<% =l("save")%>" /><%if isNumeriek(guestbook.iID) then%><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>" /><br /><br /><%end if%></td></tr></table></form><%Response.Flush 
if convertGetal(guestbook.iId)<>0 then
dim fsearch
set fsearch=new cls_fullSearch
fsearch.pattern="\[QS_GUESTBOOK:+("&guestbook.sCode&")+[\]]"
dim result
result=fsearch.search()
if not isLeeg(result) then
Response.Write "<p align=center>"&l("whereguestbookused")&"</p>"
Response.Write result
end if
set fsearch=nothing
end if%><!-- #include file="bs_gbBack.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
