<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bIntranetContacts%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><!-- #include file="includes/urlenCodeJS.asp"--><%=getBOHeader(btn_Intranet)%><%=getBOHeaderIntranet(btn_contacts)%><%dim hBody,hSubject,hBodyBGColor
hBodyBGColor="#FFFFFF"
dim contactFields, field
set contactFields=customer.contactFields(false)
if convertBool(Request.Form ("postBack")) then
if convertGetal(Request.form("resendMailID"))<>0 then
dim getTemplateMail
set getTemplateMail=new cls_mail
getTemplateMail.pick(Request.form("resendMailID"))
hBody=getTemplateMail.sBody
hBodyBGColor=getTemplateMail.sBodyBGColor
hSubject=getTemplateMail.sSubject
set getTemplateMail=nothing
else
hBody	= Request.Form ("sBody")
hSubject	= Request.Form ("sSubject")
hBodyBGColor	= Request.Form ("sBodyBGColor")
'response.write "hBody: " & hBody
end if
end if
if convertBool(Request.Form("sendBack")) then
if isLeeg(hSubject) then
message.AddError("err_mandatory")
end if
if isLeeg(removehtml(Request.Form ("sBody"))) then
message.AddError("err_mandatory")
end if
if not message.hasErrors then
sOWBodyBGColor=hBodyBGColor%><!-- #include file="bs_massMailing2.asp"--><%end if
end if
dim contactCount%><form onsubmit="javascript:if(confirm('<%=quotrepJS(l("areyousure"))%>')){document.mainform.dummy.disabled=true;document.mainform.dummy.value='<%=quotrepJS(l("pleasewait"))%>';document.mainform.sendBack.value='<%=true%>';return true;} else {return false;}" action="bs_contactSelectionActions.asp" method="post" name="mainform"><%=QS_secCodeHidden%><input type=hidden name=postBack value="<%=true%>" /><input type=hidden name=sendBack value="<%=false%>" /><%contactCount=0
massHidden()%><table align=center><tr><td class=QSlabel>N&#176; <%=l("contacts")%>:</td><td><b><%=contactCount%></b></td></tr><tr><td class=QSlabel><%=l("resend")%>:</td><td><select onchange="javascript:if(!this.value==''){if(confirm('<%=quotrepJS(l("warningresend"))%>')){document.mainform.sendBack.value='<%=false%>';document.mainform.submit()};};" name='resendMailID' style="width:650"><option></option><%dim mails, mail
set mails=customer.mails
for each mail in mails
Response.Write "<option value='" & mail & "'>" & convertEuroDate (mails(mail).dDateSent) & " - " & quotrep(mails(mail).ssubject) & "</option>"
next%></select></td></tr><tr><td class=QSlabel><%=l("subject")%>:*</td><td><input type=text size=90 name="sSubject" value="<%=quotRep(hSubject)%>" /></td></tr><tr><td class=QSlabel><%=l("message")%>:*</td><td><%createFCKInstance hBody,"siteBuilderMail","sBody"%></td></tr><tr><td class="QSlabel">Body background-color:*</td><td><input type="text" id="hBodyBGColor" name="sBodyBGColor" value="<%=quotrep(hBodyBGColor)%>" /><%=JQColorPicker("hBodyBGColor")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><%=l("sendcopyeach")%><br /><input type=text size=90 name="ccEmails" value="<%=quotRep(Request.Form ("ccEmails"))%>" /></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input  class="art-button" type="submit" name="dummy" value="<%=l("send")%>" /></td></tr></table></form><p align=center><input type=text  value="[<%=l("email")%>]" size=50 onclick="javascript:this.select();" onblur="javascript:this.value='[<%=l("email")%>]'" id=text1 name=text1 /><br /><%if contactFields.count>0 then%><%for each field in contactFields
if contactFields(field).sType<>sb_checkbox then%><input type=text  value="[<%=quotRep(contactFields(field).sFieldName)%>]" size=45 onclick="javascript:this.select();" onblur="javascript:this.value='[<%=contactFields(field).sFieldName%>]'" /><br /><%end if
next%><%end if%></p><%function massUpdate(NulOfEen)
end function
function massHidden()
dim contactKey
for each contactKey in Request.Form ("iContactIDM")
contactCount=contactCount+1
Response.Write "<input type='hidden' name='iContactIDM' value='"& contactKey &"' />" & vbcrlf
next
end function%><!-- #include file="bs_backContact.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
