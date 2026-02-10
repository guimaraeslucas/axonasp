<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><!-- #include file="bs_process.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Home)%><%dim pobj
set pobj=new cls_page
pobj.pick(decrypt(Request("iId")))
select case Request.Form ("btnValidate")
case "Save and Validate New Title"
pobj.sTitle=Request.Form ("sTitleToBeValidated")
pobj.sTitleToBeValidated=null
sendMailToContact(pobj.iUpdatedBy)
if pobj.save() then
message.Add ("fb_saveOK")
end if
case "Remove New Title"
pobj.sTitleToBeValidated=null
if pobj.save() then
message.Add ("fb_saveOK")
end if
case "Save and Validate New Text"
pobj.sValue=Request.Form ("sValueToBeValidated")
pobj.sValueToBeValidated=null
sendMailToContact(pobj.iUpdatedBy)
if pobj.save() then
message.Add ("fb_saveOK")
end if
case "Remove New Text"
pobj.sValueToBeValidated=null
if pobj.save() then
message.Add ("fb_saveOK")
end if
end select
function sendMailToContact(contactID)
dim nContact
set nContact=new cls_contact
nContact.pick(contactID)
dim notifMail
set notifMail=new cls_mail_message
notifMail.receiver=nContact.sEmail
notifMail.subject=customer.sname & " - Changes to Page '" & pobj.sTitle & "' have been validated!"
notifMail.body="<p>The webmaster has validated the changes you made to the page <b><a href='default.asp?iId=" & encrypt(pobj.iId) & "'>" & pobj.sTitle & "</a></b> on <a href='"&customer.sUrl&"'>" & customer.sUrl & "</a>.</p><p>Thank you for updating the page!</p>"
notifMail.send
set notifMail=nothing 
set nContact=nothing
end function%><p align="center">Page <b><%=pobj.sTitle%></b> was updated by <b><%=pobj.updater.sNickName%><b> on <%=convertEuroDateTime(pobj.dUpdatedOn)%></p><% 
if pobj.sTitleToBeValidated<>"" then%><form action="bs_validatepage.asp" method="post" name="validateTITLE"><input type="hidden" name="iId" value="<%=encrypt(pobj.iId)%>" /><table cellpadding="3" cellspacing="0" border="1" align="center" style="width:95%"><tr><td colspan=2 style="color:white;background-color:#666666"><b>New title to be validated!</b></td></tr><tr><td class=QSlabel><b>Current Title:</b></td><td><%=pobj.sTitle%></td></tr><tr><td class=QSlabel><b>New Title:</b></td><td><input type="text" size="60" maxlength="255" name="sTitleToBeValidated" value="<%=quotrep(pobj.sTitleToBeValidated)%>" /></td></tr><tr><td class=labe>&nbsp;</td><td><input class="art-button" type=submit value="Save and Validate New Title" name="btnValidate" onclick="javascript:return confirm('Are you sure?');" /> 
<input class="art-button" type=submit value="Remove New Title" name="btnValidate" onclick="javascript:return confirm('Are you sure?\nThis will cancel the changes by the user!');" /></td></tr></table></form><%end if%><%if pobj.sValueToBeValidated<>"" then%><form action="bs_validatepage.asp" method="post" name="validateTEXT"><input type="hidden" name="iId" value="<%=encrypt(pobj.iId)%>" /><table cellpadding="3" cellspacing="0"  border="1"  align="center" style="width:95%"><tr><td colspan=2 style="color:white;background-color:#666666"><b>New text to be validated!</b></td></tr><tr><td class=QSlabel valign="top"><b>Current text:</b></td><td><div style="width:660px;height:200px;overflow:scroll"><%=pobj.sValue%></div></td></tr><tr><td class=QSlabel valign="top"><b>New Text:</b></td><td><%createFCKInstance pobj.sValueToBeValidated,"siteBuilderMail","sValueToBeValidated"%></td></tr><tr><td class=labe>&nbsp;</td><td><input class="art-button" type=submit value="Save and Validate New Text" name="btnValidate" onclick="javascript:return confirm('Are you sure?');" /> 
<input class="art-button" type=submit value="Remove New Text" name="btnValidate" onclick="javascript:return confirm('Are you sure?\nThis will cancel the changes by the user!');" /></td></tr></table></form><%end if%><table align=center><tr><td align=center>-> <b><a href="bs_validatepages.asp"><%=l("back")%></a></b> <-</td></tr></table><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
