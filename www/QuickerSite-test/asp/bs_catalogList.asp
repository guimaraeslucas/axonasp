<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bCatalog%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Catalog)%><%dim cf
set cf=new cls_catalogField
select case request("btnaction")
case "MoveUP"
checkCSRF()
cf.moveUp()
case "MoveDOWN"
checkCSRF()
cf.moveDown()
end select
dim catalogs
set catalogs=customer.catalogs%><p align=center><%=getArtLink("bs_catalogEdit.asp",l("newcatalog"),"","","")%></p><%if catalogs.count>0 then
dim key, fields, fieldKey, fileTypes, fileKey
for each key in catalogs
set fields	= catalogs(key).fields("")
set fileTypes	= catalogs(key).fileTypes%><table align=center width='700' border=1 cellpadding=5 cellspacing=0><tr><td class=header colspan=4><a href="bs_catalogEdit.asp?iCatalogID=<%=encrypt(key)%>"><%=catalogs(key).sName%></a> <i>(iId: <%=key%>)</i><%if fields.count>0 then%>&nbsp;-&nbsp;<a href="bs_catalogItemEdit.asp?iCatalogID=<%=encrypt(key)%>"><%=l("newitem")%></a>&nbsp;-&nbsp;<a href="bs_catalogItemSearch.asp?iCatalogID=<%=encrypt(key)%>"><b><%=l("list")%></b></a><%end if%></td></tr><tr><td width=400><%=l("fields")%></td><td><%=l("attachments")%></td></tr><tr><td valign=top><%if fields.count>0 then%><table border=0 cellpadding=1 cellspacing=1><%for each fieldKey in fields%><tr><td>&bull;&nbsp;<a href="bs_catalogFieldEdit.asp?iFieldID=<%=encrypt(fieldKey)%>"><%=fields(fieldKey).sName%></a></td><td><table cellpadding=1 cellspacing=0><tr><td><%=getIcon(l("up"),"up","bs_catalogList.asp?"&QS_secCodeURL&"&amp;iFieldID="&encrypt(fieldKey)&"&amp;btnaction=MoveUP","","up"&fieldKey)%></td><td><%=getIcon(l("down"),"down","bs_catalogList.asp?"&QS_secCodeURL&"&amp;iFieldID="&encrypt(fieldKey)&"&amp;btnaction=MoveDOWN","","down"&fieldKey)%></td></tr></table></td><td><i>(iId: <%=fieldKey%>)</i></td></tr><%next%></table><br /><%end if%>-> <a href="bs_catalogFieldEdit.asp?iCatalogID=<%=encrypt(key)%>"><%=l("newfield")%></a></td><td valign=top><%if fileTypes.count>0 then%><table border=0 cellpadding=1 cellspacing=1><%for each fileKey in fileTypes%><tr><td>&bull;&nbsp;<a href="bs_catalogFileTypeEdit.asp?iFileTypeID=<%=encrypt(fileKey)%>"><%=fileTypes(fileKey).sName%></a></td><td><i>(iId: <%=fileKey%>)</i></td></tr><%next%></table><br /><%end if%>-> <a href="bs_catalogFileTypeEdit.asp?iCatalogID=<%=encrypt(key)%>"><%=l("newtype")%></a></td></tr></table><br /><%set fileTypes	= nothing
set fields	= nothing
next
else%><p align=center><%=l("nocatalogs")%></p><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
