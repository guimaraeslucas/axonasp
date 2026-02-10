<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bCatalog%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Catalog)%><%dim catalogFileType
set catalogFileType=new cls_catalogFileType
select case Request.Form ("btnaction")
case l("save")
checkCSRF()
catalogFileType.getRequestValues()
if catalogFileType.save then Response.Redirect ("bs_catalogList.asp")
case l("delete")
checkCSRF()
catalogFileType.remove()
Response.Redirect ("bs_catalogList.asp")
end select%><!-- #include file="bs_backCatalog.asp"--><form action="bs_catalogFileTypeEdit.asp" method="post" name="mainform"><%=QS_secCodeHidden%><input type="hidden" name="iFileTypeID" value="<%=encrypt(catalogFileType.iID)%>" /><input type="hidden" name="iCatalogID" value="<%=encrypt(catalogFileType.iCatalogID)%>" /><table align=center cellpadding="2"><tr><td colspan=2 class=header><%=l("general")%>:</td></tr><tr><td class=QSlabel><%=l("catalog")%>:</td><td><%=catalogFileType.catalog.sName%></td></tr><tr><td class=QSlabel><%=l("name")%>:*</td><td><input type=text size=40 name="sName" value="<%=quotRep(catalogFileType.sName)%>" /></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=btnaction value="<% =l("save")%>" /><%if isNumeriek(catalogFileType.iID) then%><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>"><br /><br /><%end if%></td></tr></table></form><!-- #include file="bs_backCatalog.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
