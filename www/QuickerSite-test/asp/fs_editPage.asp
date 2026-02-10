<!-- #include file="begin.asp"-->


<%pagetoemail=true
editPage=true
dim EPSuccess, EPSaveSuccess
EPSuccess=false
EPSaveSuccess=false
dim accessToEP
accessToEP=false
dim getTperP,getBperP
set getTperP=logon.contact.getTper
set getBperP=logon.contact.getBper
logon.contact.createUserFilesFolder()
if convertBool(Request.Form ("postback")) then
checkCSRF()
if getTPerP.exists(selectedPage.iId) then
if not isLeeg(customer.sNotifValidate) then
selectedPage.sTitleToBeValidated	= convertStr(Request.Form ("sTitle"))
else
selectedPage.sTitle	= convertStr(Request.Form ("sTitle"))
end if
end if 
if getBPerP.exists(selectedPage.iId) then
if not isLeeg(customer.sNotifValidate) then
selectedPage.sValueToBeValidated = convertStr(Request.Form ("sValue"))
else
selectedPage.sValue	= convertStr(Request.Form ("sValue"))
end if
selectedPage.sExternalURLPrefix	= convertStr(Request.Form ("sExternalURLPrefix"))
selectedPage.sExternalURL	= convertStr(Request.Form ("sExternalURL"))
selectedPage.bOpenInNewWindow	= convertBool(Request.Form ("bOpenInNewWindow"))
end if
if getTPerP.exists(selectedPage.iId) or getBPerP.exists(selectedPage.iId) then
if selectedPage.Save then
message.Add("fb_saveOK")
EPSaveSuccess=true
if not isLeeg(customer.sNotifValidate) then
if convertStr(selectedPage.sTitleToBeValidated)<>convertStr(selectedPage.sTitle) or convertStr(selectedPage.sValueToBeValidated)<>convertStr(selectedPage.sValue) then
selectedPage.iUpdatedBy=logon.contact.iId
selectedPage.dUpdatedOn=now()
selectedPage.Save()
dim notifMail
set notifMail=new cls_mail_message
notifMail.receiver=customer.sNotifValidate
notifMail.subject=customer.sname & " - Page to be validated!"
notifMail.body="<p>"& logon.contact.sNickname&" ("&logon.contact.sEmail&") has updated the page <b>" & selectedPage.sTitle & "</b> on <a href='"&customer.sUrl&"'>" & customer.sUrl & "</a>.</p><p>This update needs to be validated by you.<p></p>Please logon to the backsite of <a href='" & customer.sUrl & "'>" & customer.surl & "</a> to validate the changes.</p>"
notifMail.send
set notifMail=nothing
end if
end if
end if
end if 
end if
if selectedPage.sTitleToBeValidated<>selectedPage.sTitle and not isLeeg(selectedPage.sTitleToBeValidated) then
selectedPage.sTitle=selectedPage.sTitleToBeValidated
end if
if selectedPage.sValueToBeValidated<>selectedPage.sValue and not isLeeg(selectedPage.sValueToBeValidated) then
selectedPage.sValue=selectedPage.sValueToBeValidated
end if%><!-- #include file="includes/commonheader.asp"--><body style="color:#000;background-color:#FFF;background-image:none" onload="javascript:window.focus();"><%'if not EPSaveSuccess then%><form method="post" action="fs_editPage.asp" name="mainform"><input type="hidden" name="iId" value="<%=encrypt(selectedPage.iId)%>" /><input type="hidden" name="postback" value="<%=true%>" /><%=QS_secCodeHidden%><table align="center" style="background-color:#FFF"><%if getTPerP.exists(selectedPage.iId) then%><tr><td class="QSlabel"><%=l("title")%>:*</td><td><input type="text" size="40" maxlength="100" name="sTitle" value="<%=quotrep(selectedPage.sTitle)%>" /></td></tr><%accessToEP=true
end if
if getBPerP.exists(selectedPage.iId) then
if not isLeeg(selectedPage.sExternalURL) then
dim urlTypeList
set urlTypeList=new cls_urlTypeList%><tr><td class=QSlabel><%=l("externalurl")%>:*</td><td><select name=sExternalURLPrefix><%=urlTypeList.showSelected("option",selectedPage.sExternalURLPrefix)%></select>&nbsp;<input size=40 maxlength=254 type=text value="<%=quotRep(selectedPage.sExternalURL)%>" name=sExternalURL /></td></tr><tr><td class=QSlabel><%=l("openinnewwindow")%></td><td><input name=bOpenInNewWindow type=checkbox value=1 <%if selectedPage.bOpenInNewWindow then Response.Write "checked"%> /></td></tr><%elseif not convertBool(selectedPage.bContainerPage) then%><tr><td colspan="2"><%createFCKInstance selectedPage.sValue,"siteBuilder","sValue"%></td></tr><%end if
accessToEP=true
end if%><tr><td class="QSlabel">&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class="QSlabel">&nbsp;</td><td><input type="submit" name="dummy" value="<% =l("save")%>" /></td></tr></table></form><%'end if%><%=message.showAlert()%><%=cPopup.dumpJS()%> 
<%if not accessToEP then Response.Redirect ("asp/noaccess.htm")%></body></html><%cleanUPASP%>
