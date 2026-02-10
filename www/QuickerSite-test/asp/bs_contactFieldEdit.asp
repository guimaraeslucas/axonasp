<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bIntranetContacts%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Intranet)%><%=getBOHeaderIntranet(btn_contacts)%><%dim contactField
set contactField=new cls_contactField
dim postBack
postBack=convertBool(Request.Form("postBack"))
if postBack then
contactField.getRequestValues()
end if
select case Request.Form ("btnaction")
case l("save")
checkCSRF()
contactField.getRequestValues()
if contactField.save() then
Response.Redirect ("bs_contactFields.asp")
end if
case l("delete")
checkCSRF()
contactField.remove()
Response.Redirect ("bs_contactFields.asp")
end select
dim fixedFieldTypeList
set fixedFieldTypeList=new cls_contactFieldTypeList%><form action="bs_contactFieldEdit.asp" method="post" name=mainform><%=QS_secCodeHidden%><input type="hidden" value="<%=true%>" name=postback /><input type="hidden" value="<% =encrypt(contactField.iID)%>" name=iFieldID /><table align=center cellpadding="2"><tr><td colspan=2 class=header><%=l("field")%>:</td></tr><tr><td class=QSlabel><%=l("name")%>:*</td><td><input type=text size=30 name="sFieldName" value="<%=quotRep(contactField.sFieldName)%>" /></td></tr><tr><td class=QSlabel>&nbsp;</td><td><%=l("emaildefaultfield")%></td></tr><tr><td class=QSlabel><%=l("type")%>:*</td><td><select name=sType onchange="javascript:document.mainform.submit();"><%=fixedFieldTypeList.showSelected("option",contactField.sType)%></select></td></tr><%if contactField.sType=sb_select then%><tr><td class=QSlabel><%=l("values")%>:*<br /><i>(<%=l("enterseplist")%>)</i></td><td><textarea name=sValues cols=50 rows=10><% =quotRep(contactField.sValues)%></textarea></td></tr><%end if%><%if contactField.sType=sb_comment then%><tr><td class=QSlabel><%=l("comment")%>:*</td><td><%createFCKInstance contactField.sValues,"siteBuilderRichText","sValues"%></td></tr><%end if%><%if contactField.sType<>sb_checkbox  and contactField.sType<>sb_comment then%><tr><td class=QSlabel><%=l("mandatory")%></td><td><input type=checkbox  name="bMandatory" value="1" <%if contactField.bMandatory then Response.Write "checked"%> /></td></tr><%end if
if contactField.sType<>sb_comment then%><tr><td class=QSlabel><%=l("searchfield")%></td><td><input type=checkbox  name="bSearchField" value="1" <%if contactField.bSearchField then Response.Write "checked"%> /></td></tr><%end if%><tr><td class=QSlabel><%=l("publicprofile")%></td><td><input type=checkbox  name="bProfile" value="1" <%if contactField.bProfile then Response.Write "checked"%> /></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=btnaction value="<% =l("save")%>" /><%if isNumeriek(contactField.iID) then%><input class="art-button" type=submit name=btnaction value="<% =l("delete")%>" onclick="javascript:return confirm('<%=l("areyousure")%> <%=l("deletefieldfb")%>')" /><%end if%></td></tr></table></form><!-- #include file="bs_backContactField.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
