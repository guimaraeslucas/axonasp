<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bIntranetContacts%><%dim contact
set contact=new cls_contact
if session("bLoadCPinPP")="yes" then
session("bLoadCPinPP")=""
response.redirect ("bs_contactHome.asp?iCPP=" & contact.iId)
end if
dim contactFields, contactField
set contactFields=customer.contactFields(false)
dim postback
postback=convertBool(Request.Form ("postback"))
if postback then
saveHiddenValues=true
contact.getRequestValues(contactFields)
contact.iStatus=convertGetal(Request.Form ("iStatus"))
end if
select case Request.Form ("btnaction")
case l("save")
checkCSRF()
contact.getRequestValues(contactFields)
contact.iStatus=convertGetal(Request.Form ("iStatus"))
if contact.save(contactFields) then
message.Add("fb_saveOK")
end if
case l("delete")
checkCSRF()
contact.delete
cPopup.close()
case l("savAndResetPw")
checkCSRF()
contact.getRequestValues(contactFields)
contact.iStatus=convertGetal(Request.Form ("iStatus"))
if contact.save(contactFields) then
contact.resetPW()
message.Add("fb_saveOK")
message.Add("fb_passwordreset")
end if
end select
dim fields
set fields=contact.fields%><!-- #include file="includes/commonheader.asp"--><body style="background-color:#FFF"><!-- #include file="bs_initBack.asp"--><form action="bs_contactEdit.asp" method="post" name=mainform><%=QS_secCodeHidden%><input type="hidden" name="iContactID" value="<%=encrypt(contact.iID)%>" /><input type="hidden" value="<%=true%>" name="postback" /><table align="center" cellpadding="2"><tr><td colspan=2 class=header><%=l("logginin")%>:</td></tr><tr><td class=QSlabel><%=l("email")%>:*</td><td><input type=email size=30 maxlength=100 name="sEmail" value="<%=quotRep(contact.sEmail)%>" /><%if not isLeeg(contact.sOrigEmail) then Response.Write "<br />Original: " & clickEmail(contact.sOrigEmail)%></td></tr><tr><td class=QSlabel><%=l("password")%>:*</td><td><input type=password size=15 maxlength=20 name="sPw" value="<%=contact.sPw%>" /> <a href="#" onclick="mainform.sPw.value='<%=GeneratePassWord%>';return false"><i><%=l("random")%></i></a></td></tr><tr><td class=QSlabel><%=l("nickname")%>:*</td><td><input type=text size=15 maxlength=30 name="sNickname" value="<%=quotRep(contact.sNickname)%>" /></td></tr><tr><td class=QSlabel><%=l("allowprivatemessages")%></td><td><input type="radio" name="bGetEmailsFromSite" <%=convertChecked(contact.bGetEmailsFromSite)%> value="<%=true%>" /> <%=l("yes")%>  <input value="<%=false%>" type="radio" name="bGetEmailsFromSite" <%=convertChecked(not contact.bGetEmailsFromSite)%> /> <%=l("no")%></td></tr><tr><td class=QSlabel width=250><%=l("memberrole")%>:</td><td><select name="iStatus" onchange="javascript:document.mainform.submit();"><%=cslist.showSelected("option",contact.iStatus)%></select></td></tr><tr><td colspan=2><hr /></td></tr><tr><td colspan=2 class=header><%=l("contactdata")%>:</td></tr><%for each contactField in contactFields
if contactFields(contactField).sType<>sb_comment then%><tr><td class=QSlabel width=160><%=contactFields(contactField).sFieldname%><%if contactFields(contactField).bMandatory then%>*<%end if%></td><td><%select case contactFields(contactField).sType
case sb_text, sb_url, sb_email%><input type=text size=30 maxlength=300 name="<%=encrypt(contactField)%>" value="<%=quotRep(fields(contactField))%>" /><%case sb_textarea%><textarea cols=29 rows=3 name="<%=encrypt(contactField)%>"><%=quotRep(fields(contactField))%></textarea><%case sb_checkbox%><input type=checkbox name="<%=encrypt(contactField)%>" value="checked" <%=fields(contactField)%> /><%case sb_select%><select name="<%=encrypt(contactField)%>"><%=contactFields(contactField).showSelected(fields(contactField))%></select><%case sb_date%><input type="text" id="<%=encrypt(contactField)%>" name="<%=encrypt(contactField)%>" value="<%=convertDateToPicker(fields(contactField))%>" /><%=JQDatePicker(encrypt(contactField))%><%case sb_richtext%><%createFCKInstance fields(contactField),"siteBuilderRichText",encrypt(contactField)%><%end select%></td></tr><%end if
next%><%if isNumeriek(contact.iID) then%><tr><td colspan=2><hr /></td></tr><tr><td class=QSlabel><%=l("createdon")%>:</td><td><%=convertEuroDateTime(contact.dCreatedTS)%></td></tr><tr><td class=QSlabel><%=l("updatedon")%>:</td><td><%=convertEuroDateTime(contact.dUpdatedTS)%></td></tr><%if not isLeeg(contact.dLastLoginTS) then%><tr><td class=QSlabel><%=l("lastlogin")%>:</td><td><%=convertEuroDateTime(contact.dLastLoginTS)%></td></tr><%end if
end if%><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=btnaction value="<% =l("save")%>" /><%if isNumeriek(contact.iID) then%><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>" /><br /><br /><%if contact.iStatus>cs_silent then%><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>&nbsp;<%=l("abouttosentpw")%>');" value="<% =l("savAndResetPw")%>" /><br /><%end if
end if%></td></tr></table></form><%if contactFields.count>0 and isNumeriek(contact.iID) then%><p align=center><a href="bs_contactEdit.asp"><b><%=l("addcontact")%></b></a></p><%end if%><!-- #include file="bs_endBack.asp"--><%=message.showAlert()%><%=cPopup.dumpJS()%> 
</body></html><%cleanUPASP%>
