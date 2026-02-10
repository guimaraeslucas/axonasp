<!-- #include file="begin.asp"-->
<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bForms%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Catalog)%><%dim FormField
set FormField=new cls_FormField
dim postBack
postBack=convertBool(Request.Form("postBack"))
if postBack then
FormField.getRequestValues()
end if
select case Request.Form ("btnaction")
case l("save")
checkCSRF()
FormField.getRequestValues()
if FormField.save then Response.Redirect ("bs_formFields.asp?iFormID=" & encrypt(FormField.iFormID))
case l("delete")
checkCSRF()
FormField.remove()
Response.Redirect ("bs_formFields.asp?iFormID=" & encrypt(FormField.iFormID))
end select
dim fixedFieldTypeList
set fixedFieldTypeList=new cls_formFieldTypeList
dim placementList
set placementList=new cls_placementList%><form action="bs_FormFieldEdit.asp" method="post" name=mainform><%=QS_secCodeHidden%><input type="hidden" name="iFormFieldID" value="<%=encrypt(FormField.iID)%>" /><input type="hidden" name="iFormID" value="<%=encrypt(FormField.iFormID)%>" /><input type="hidden" value="<%=true%>" name="postback" /><table align="center"><tr><td colspan=2 class="header"><%=l("general")%>:</td></tr><tr><td class=QSlabel><%=l("form")%>:</td><td><%=FormField.form.sName%></td></tr>
<tr><td class=QSlabel><%=l("name")%>:*</td><td><input type=text size="40" name="sName" value="<%=sanitize(FormField.sName)%>" /></td></tr>
<%
select case FormField.sType
	case sb_ff_text, sb_ff_url, sb_ff_email,sb_ff_textarea
%>
<tr><td class=QSlabel>Placeholder:</td><td><input type=text size="40" name="sPlaceholder" value="<%=sanitize(FormField.sPlaceholder)%>" /></td></tr>
<%
end select

dim totalRang
if convertGetal(FormField.iId)=0 then
totalRang=FormField.form.fields.count+1
FormField.iRang=totalRang
else
totalRang=FormField.form.fields.count
end if%><tr><td class=QSlabel><%=l("sortorder")%>:*</td><td><select name="iRang"><%=numberList(1,totalRang,1,FormField.iRang)%></select></td></tr><tr><td class=QSlabel><%=l("type")%>:*</td><td><select name=sType onchange="javascript:document.mainform.submit();"><%=fixedFieldTypeList.showSelected("option",FormField.sType)%></select></td></tr><%if FormField.sType=sb_ff_hidden then%><tr><td class=QSlabel><%=l("defaultvalue")%>:</td><td><input type="text" value="<%=sanitize(FormField.sValues)%>" name="sValues" size="30" maxlength="250" /></td></tr><%end if%><%if FormField.sType=sb_ff_textarea then%><tr><td class=QSlabel><%=l("columns")%>:</td><td><select name="iCols"><% =numberList(10,100,1,FormField.iCols)%></select></td></tr><tr><td class=QSlabel><%=l("rows")%>:</td><td><select name="iRows"><% =numberList(2,50,1,FormField.iRows)%></select></td></tr><tr><td class=QSlabel><%=l("maxlength")%>:</td><td><select name="iMaxlength"><% =numberList(0,2500,50,FormField.iMaxlength)%></select>&nbsp;<%=l("characters")%> - 0: <%=l("nolimit")%></td></tr><%end if%><%if FormField.sType=sb_ff_text or FormField.sType=sb_ff_url or FormField.sType=sb_ff_email then%><tr><td class=QSlabel><%=l("size")%>:</td><td><select name="iSize"><% =numberList(2,100,1,FormField.iSize)%></select></td></tr><tr><td class=QSlabel><%=l("maxlength")%>:</td><td><select name="iMaxlength"><% =numberList(1,1024,1,FormField.iMaxlength)%></select></td></tr><%end if%><%if FormField.sType=sb_ff_select or FormField.sType=sb_ff_radio then%><tr><td class=QSlabel><%=l("values")%>:*<br /><i>(<%=l("enterseplist")%>)</i></td><td><textarea name=sValues cols=50 rows=10><% =quotRep(FormField.sValues)%></textarea></td></tr><%if FormField.sType=sb_ff_radio then%><tr><td class=QSlabel>Allow multiple selections?</td><td><input type=checkbox name="bAllowMS" value="1" <%if FormField.bAllowMS then Response.Write "checked"%> /></td></tr><%end if
end if%><%if FormField.sType=sb_ff_comment then%><tr><td class=QSlabel><%=l("comment")%>:*</td><td><%createFCKInstance FormField.sValues,"siteBuilderRichText","sValues"%></td></tr><%end if%><%if FormField.sType=sb_ff_radio then%><tr><td class=QSlabel><%=l("placement")%>:</td><td><select name=sRadioPlacement><%=placementList.showSelected("option",FormField.sRadioPlacement)%></select></td></tr><%end if%><%if FormField.sType=sb_ff_file or FormField.sType=sb_ff_image then%><tr><td class=QSlabel><%=l("storeinfolder")%>:</td><td><input type="text" size="65" name="sFileLocation" maxlength="255" value="<%=quotrep(FormField.sFileLocation)%>" /><br /><span style="font-size:0.8em">default: <%=server.MapPath (application("QS_CMS_userfiles"))%></span></td></tr><%if FormField.sType=sb_ff_file then%><tr><td class=QSlabel><%=l("allowedfiletypes")%>:</td><td><input type="text" size="65" name="sAllowedExtensions" maxlength="255" value="<%=quotrep(FormField.sAllowedExtensions)%>"><br /><span style="font-size:0.8em"><%=l("cs_fileextx")%> - default: <%=l("nolimit")%></span></td></tr><%end if%><tr><td class=QSlabel><%=l("maxfilesize")%>:</td><td><select name=iMaxFileSize><% =numberList(16,4096,16,FormField.iMaxFileSize)%></select> kB</td></tr><%end if%><%if FormField.sType=sb_ff_email and FormField.form.bAutoResponder then%><tr><td class=QSlabel><%=l("enableautoresponse")%></td><td><input type=checkbox name="bAutoResponder" value="1" <%if FormField.bAutoResponder then Response.Write "checked"%> /></td></tr><%end if%><%if FormField.sType=sb_ff_email and FormField.form.bSendEmail then%><tr><td class=QSlabel><%=l("useasenderemail")%></td><td><input type=checkbox name="bUseForSending" value="1" <%if FormField.bUseForSending then Response.Write "checked"%> /></td></tr><%end if%><%if FormField.sType<>sb_ff_comment and FormField.sType<>sb_ff_hidden then%><tr><td class=QSlabel><%=l("mandatory")%></td><td><input type=checkbox name="bMandatory" value="1" <%if FormField.bMandatory then Response.Write "checked"%> /></td></tr><%end if%><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=btnaction value="<% =l("save")%>" /><%if isNumeriek(FormField.iID) then%><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>" /><br /><br /><%end if%></td></tr></table></form><!-- #include file="bs_backFormField.asp"--><!-- #include file="bs_FormExpl.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
