<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bCatalog%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Catalog)%><%dim catalog
set catalog=new cls_catalog
dim postBack
postBack=convertBool(Request.Form("postBack"))
if postBack then
catalog.getRequestValues()
end if
select case Request.Form ("btnaction")
case l("save")
checkCSRF()
catalog.getRequestValues()
if catalog.save then Response.Redirect ("bs_catalogEdit.asp?iCatalogID=" & encrypt(catalog.iId) & "&fbMessage=fb_saveOK")
case l("delete")
checkCSRF()
catalog.remove
Response.Redirect ("bs_catalogList.asp")
end select
dim fe
set fe=new cls_fileexplorer
dim catalogOrderByList
set catalogOrderByList=new cls_catalogOrderByList
dim fields
set fields=catalog.fields("public")
dim filetypes
set filetypes=catalog.filetypes%><%if fields.count>0 then%><p align="center"><%=getArtLink("bs_catalogItemEdit.asp?iCatalogID=" & encrypt(catalog.iID),l("newitem"),"","","")%>&nbsp;&nbsp;
<%=getArtLink("bs_catalogItemSearch.asp?iCatalogID=" & encrypt(catalog.iID),l("list"),"","","")%></p><%elseif convertGetal(catalog.iID)<>0 then%><p align="center"><%=getArtLink("bs_catalogFieldEdit.asp?iCatalogID=" & encrypt(catalog.iID),l("newfield"),"","","")%></p><%end if%><form action="bs_catalogEdit.asp" method="post" name="mainform"><%=QS_secCodeHidden%><input type="hidden" name="iCatalogID" value="<%=encrypt(catalog.iID)%>" /><input type="hidden" value="<%=true%>" name="postback" /><table align="center" cellpadding="2"><tr><td colspan=2 class="header"><%=l("general")%>:</td></tr><tr><td class=QSlabel><%=l("name")%>:*</td><td><input type=text size=30 maxlength=255 name="sName" value="<%=quotRep(catalog.sName)%>"></td></tr><tr><td class=QSlabel><%=l("name")%>&nbsp;<%=l("item")%>:*</td><td><input type=text size=30 maxlength=50 name="sItemName" value="<%=quotRep(catalog.sItemName)%>"></td></tr><tr><td class=QSlabel><%=l("online")%></td><td><input type=checkbox name="bOnline" value="1" <%=convertChecked(catalog.bOnline)%>></td></tr><tr><td class=QSlabel>Push RSS</td><td><input type=checkbox name="bPushRSS" value="1" <%=convertChecked(catalog.bPushRSS)%>></td></tr><tr><td class=QSlabel><%=l("searchable")%></td><td><input type=checkbox name="bSearchable" value="1" <%=convertChecked(catalog.bSearchable)%>></td></tr><%if QS_ASPX then%><tr><td class=QSlabel><%=l("autothumb")%></td><td><input  onclick="javascript:document.mainform.submit();" type=checkbox name="bAutoThumb" value="1" <%=convertChecked(catalog.bAutoThumb)%>></td></tr><%if catalog.bAutoThumb then%><tr><td class=QSlabel><%=l("thumbsize")%></td><td><select name=iMaxThumbSize><%=numberList(10,1000,10,catalog.iMaxThumbSize)%></select> px</td></tr><tr><td class=QSlabel><%=l("resizepicto")%></td><td><select name=sResizePicTo><%=numberList(480,2800,16,catalog.sResizePicTo)%></select> px</td></tr><tr><td class=QSlabel><%=l("expl_labelfullimage")%></td><td><input type="" size=40 maxlength="200" name="sFullImage" value="<%=sanitize(catalog.sFullImage)%>"></td></tr><%end if
end if%><tr><td class=QSlabel><%=l("useshadow")%></td><td><input type=checkbox name="bUseShadow" value="1" <%=convertChecked(catalog.bUseShadow)%>></td></tr><tr><td class=QSlabel><%=l("sortitemsby")%>:</td><td><select name=sOrderItemsBy><%=catalogOrderByList.showSelected("option", catalog.sOrderItemsBy)%></select></td></tr><tr><td class=QSlabel><%=l("pagesize")%>:</td><td><select name=iPageSize><%=numberList(1,100,1,catalog.iPageSize)%></select></td></tr><%if customer.forms.count>0 then%><tr><td class=QSlabel><%=l("form")%>:</td><td><select name=iFormID onchange="javascript:mainform.submit();"><option value="">&nbsp;</option><%=customer.showSelectedForm("option", catalog.iFormID)%></select></td></tr><%end if%>

<tr>
	<td class="QSlabel">Path to attachments:</td>
	<td><select name="sFilePath"><%=fe.SelectBoxFolders(server.MapPath(C_VIRT_DIR & Application("QS_CMS_userfiles")),catalog.sFilePath)%></select></td>
</tr>


<%if isNumeriek(catalog.iFormID) then%><tr><td class=QSlabel><%=l("textlink")%>:*<br />('<%=l("register")%>', '<%=l("order")%>',...)</td><td><input type=text size=30 name="sFormTitle" value="<%=quotRep(catalog.sFormTitle)%>"></td></tr><%end if%><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=btnaction value="<% =l("save")%>" /><%if isNumeriek(catalog.iID) then%><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>"><br /><br /><%end if%></td></tr><tr><td class="QSlabel" width="250" valign="top"><input type="text" onclick="javascript:this.select();" value="{ITEMENCID}" / id=text1 name=text1><br /><input type="text" onclick="javascript:this.select();" value="{ITEMTITLE}" / id=text1 name=text1><br /><input type="text" onclick="javascript:this.select();" value="{ITEMDATE}" / id=text1 name=text1><br /><input type="text" onclick="javascript:this.select();" value="{ITEMPICTURE}" / id=text1 name=text1>	<br /><%dim f
for each f in fields%><input type="text" onclick="javascript:this.select();" value="{<%=ucase(fields(f).sName)%>}" /><br /><%next
for each f in filetypes%><input type="text" onclick="javascript:this.select();" value="{<%=ucase(filetypes(f).sName)%>}" id=text2 name=text2><br /><%next%></td><td>ItemVIEW: 
<br /><textarea cols="60" rows="7" name="sItemView"><%=sanitize(catalog.sItemView)%></textarea><br />RSSView1: <i><%=C_DIRECTORY_QUICKERSITE%>/rss.asp?iId=XXXXXX&viewtype=1</i><br /><textarea cols="60" rows="7" name="sRSSView1"><%=sanitize(catalog.sRSSView1)%></textarea><br />RSSView2: <i><%=C_DIRECTORY_QUICKERSITE%>/rss.asp?iId=XXXXXX&viewtype=2</i><br /><textarea cols="60" rows="7" name="sRSSView2"><%=sanitize(catalog.sRSSView2)%></textarea><br />RSSView3: <i><%=C_DIRECTORY_QUICKERSITE%>/rss.asp?iId=XXXXXX&viewtype=3</i><br /><textarea cols="60" rows="7" name="sRSSView3"><%=sanitize(catalog.sRSSView3)%></textarea></td></tr></table></form><!-- #include file="bs_backCatalog.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
