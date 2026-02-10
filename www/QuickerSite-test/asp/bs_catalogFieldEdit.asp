<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bCatalog%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Catalog)%><%dim catalogField
set catalogField=new cls_catalogField
dim postBack
postBack=convertBool(Request.Form("postBack"))
if postBack then
catalogField.getRequestValues()
end if
select case Request.Form ("btnaction")
case l("save")
checkCSRF()
catalogField.getRequestValues()
if catalogField.save then Response.Redirect ("bs_catalogList.asp")
case l("delete")
checkCSRF()
catalogField.remove()
Response.Redirect ("bs_catalogList.asp")
end select
dim fixedFieldTypeList
set fixedFieldTypeList=new cls_fixedFieldTypeList%><!-- #include file="bs_backCatalog.asp"--><form action="bs_catalogFieldEdit.asp" method="post" name=mainform><%=QS_secCodeHidden%><input type=hidden name=iFieldID value="<%=encrypt(catalogField.iID)%>" /><input type=hidden name=iCatalogID value="<%=encrypt(catalogField.iCatalogID)%>" /><input type="hidden" value="<%=true%>" name=postback /><table align=center cellpadding="2"><tr><td colspan=2 class=header><%=l("general")%>:</td></tr><tr><td class=QSlabel><%=l("catalog")%>:</td><td><%=catalogField.catalog.sName%></td></tr><tr><td class=QSlabel><%=l("name")%>:*</td><td><input type=text size=40 name="sName" value="<%=quotRep(catalogField.sName)%>" /></td></tr><tr><td class=QSlabel><%=l("type")%>:*</td><td><select name=sType onchange="javascript:document.mainform.submit();"><%=fixedFieldTypeList.showSelected("option",catalogField.sType)%></select></td></tr><%if catalogField.sType=sb_select then%><tr><td class=QSlabel><%=l("values")%>:*<br /><i>(<%=l("enterseplist")%>)</i></td><td><textarea name=sValues cols=50 rows=10><% =quotRep(catalogField.sValues)%></textarea></td></tr><%end if%><tr><td class=QSlabel><%=l("publicfield")%></td><td><input type=checkbox  name="bPublic" value="1" <%if catalogField.bPublic then Response.Write "checked"%> /></td></tr><%if catalogField.sType<>sb_checkbox  then%><tr><td class=QSlabel><%=l("mandatory")%></td><td><input type=checkbox  name="bMandatory" value="1" <%if catalogField.bMandatory then Response.Write "checked"%> /></td></tr><%end if%><tr><td class=QSlabel><%=l("searchfield")%></td><td><input type=checkbox  name="bSearchField" value="1" <%if catalogField.bSearchField then Response.Write "checked"%> /></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=btnaction value="<% =l("save")%>" /><%if isNumeriek(catalogField.iID) and catalogField.catalog.fields("").count>1 then%><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>" /><br /><br /><%end if%></td></tr></table></form><!-- #include file="bs_backCatalog.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
