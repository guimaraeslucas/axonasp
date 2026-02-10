<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bCatalog%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Catalog)%><%dim itemSearch
set itemSearch=new cls_itemSearch
dim catalogFields
set catalogFields=itemSearch.catalog.fields("search")
set itemSearch.searchFields=catalogFields
itemSearch.getRequestValues()
dim inputFields
set inputFields=itemSearch.inputFields
dim dateFields
set dateFields=itemSearch.dateFields
dim lowerGreaterThanList
set lowerGreaterThanList=new cls_lowerGreaterThanList
dim catalogItem
set catalogItem=new cls_catalogItem
if isNumeriek(catalogItem.iId) then
catalogItem.copy() 
Response.Redirect ("bs_catalogItemSearch.asp?iCatalogID="&encrypt(itemSearch.iCatalogId)&"#"&encrypt(catalogItem.iId))
end if%><!-- #include file="bs_backCatalog.asp"--><form name="mainform" action="bs_catalogItemSearch.asp" method=post><table cellpadding="2" align=center style="margin-top:10px"><tr><td class=QSlabel><%=l("catalog")%>:</td><td><select name=iCatalogID onchange="javascript:document.mainform.submit();"><option value=""><%=l("select")%>...</option><%=customer.showSelectedCatalog("option", itemSearch.iCatalogID)%></select></td></tr><%if isNumeriek(itemSearch.iCatalogID) then%><tr><td class=QSlabel><%=l("title")%>:</td><td><input type=text name=sTitle size=15 value="<%=quotRep(itemSearch.sTitle)%>" /></td></tr><%dim catalogField
for each catalogField in catalogFields%><tr><td class=QSlabel><%=catalogFields(catalogField).sName%></td><td><%select case catalogFields(catalogField).sType
case sb_text,sb_textarea,sb_url,sb_email,sb_richtext%><input type=text size=15 name="<%=encrypt(catalogField)%>" value="<%=quotRep(inputFields(catalogField))%>" /><%case sb_checkbox%><input type=checkbox name="<%=encrypt(catalogField)%>" value="checked" <%=inputFields(catalogField)%> /><%case sb_select%><select name="<%=encrypt(catalogField)%>"><%=catalogFields(catalogField).showSelected(inputFields(catalogField))%></select><%case sb_date%><i><%=l("between")%></i> 
 <input type="text" id="from<%=encrypt(catalogField)%>" name="from<%=encrypt(catalogField)%>" value="<%=dateFields("from"&catalogField)%>" /><%=JQDatePicker("from" & encrypt(catalogField))%> <i><%=l("and")%></i> 
 <input type="text" id="untill<%=encrypt(catalogField)%>" name="untill<%=encrypt(catalogField)%>" value="<%=dateFields("untill"&catalogField)%>" /><%=JQDatePicker("untill" & encrypt(catalogField))%><%end select%></td></tr><%next%><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit value="<%=l("search")%>" name=btnaction />&nbsp;<input class="art-button" type=submit value="<%=l("excel")%>" name=btnaction /></td></tr><%end if%></table></form><%dim addNew
addnew=false
if isNumeriek(itemSearch.iCatalogID) and itemSearch.catalog.fields("").count>0 then
addNew=true%><br /><table align=center><tr><td align=center><%=getArtLink("bs_catalogItemEdit.asp?iCatalogID=" & encrypt(itemSearch.iCatalogID),l("newitem") & "&nbsp;" & l("for") & "&nbsp;" & itemSearch.catalog.sName,"","","")%></td></tr></table><%end if%><%if isNumeriek(itemSearch.iCatalogID) then
Response.write itemSearch.resultTable
end if%><br /><table align=center cellpadding=4 cellspacing=0 border=0><tr><td class="<%=convertTF(true)%>" style="<%=convertTFS(true)%>" width=15 align=center valign=middle>V</td><td style="<%=convertTFS(true)%>"><%=l("online")%></td><td>&nbsp;</td><td style="<%=convertTFS(false)%>" class="<%=convertTF(false)%>" width=15 align=center valign=middle>X</td><td style="<%=convertTFS(false)%>"><%=l("offline")%></td></tr></table><br /><%if addNew then%><table align=center><tr><td align=center><%=getArtLink("bs_catalogItemEdit.asp?iCatalogID=" & encrypt(itemSearch.iCatalogID),l("newitem") & "&nbsp;" & l("for") & "&nbsp;" & itemSearch.catalog.sName,"","","")%></td></tr></table><br /><%end if%><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
