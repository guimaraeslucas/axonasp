<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bCatalog%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Catalog)%><%dim catalogItem
set catalogItem=new cls_catalogItem
dim postBack
postBack=convertBool(Request.Form("postBack"))
if postBack then
catalogItem.getRequestValues()
end if
if not isLeeg(Request.querystring("iFileID")) then
checkCSRF()
catalogItem.removeFile(decrypt(Request.querystring("iFileID")))
end if
select case Request.Form ("btnaction")
case l("save")
checkCSRF()
catalogItem.getRequestValues()
if catalogItem.save then 
Response.Redirect ("bs_catalogItemSearch.asp?iCatalogID=" & encrypt(catalogItem.iCatalogID) & "#" & encrypt(catalogItem.iID))
end if
case l("delete")
checkCSRF()
catalogItem.remove()
Response.Redirect ("bs_catalogItemSearch.asp?iCatalogID=" & encrypt(catalogItem.iCatalogID))
case l("uploadattachment")
checkCSRF()
catalogItem.getRequestValues()
if catalogItem.save then 
Response.Redirect ("bs_catalogItemFile.asp?iItemID="& encrypt(catalogItem.iID))
end if
case l("uploadpic")
checkCSRF()
catalogItem.getRequestValues()
if catalogItem.save then 
Response.Redirect ("bs_catalogItemPic.asp?iItemID="& encrypt(catalogItem.iID))
end if
case l("deletepic")
checkCSRF()
catalogItem.getRequestValues()
catalogItem.removePic()
catalogItem.save
end select
dim catalogFields, catalogField
set catalogFields=catalogItem.catalog.fields("")
dim fields
set fields=catalogItem.fields
dim fileTypes
set fileTypes=catalogItem.catalog.fileTypes
dim files
set files=catalogItem.files%><table align=center><tr><td align=center><%=getArtLink("bs_catalogItemSearch.asp?iCatalogID=" & encrypt(catalogItem.iCatalogID),l("back"),"","","")%></td></tr></table><br /><form method="post" action="bs_catalogItemEdit.asp" name=mainform><%=QS_secCodeHidden%><input type=hidden name=iItemID value="<%=encrypt(catalogItem.iID)%>" /><input type=hidden name=iCatalogID value="<%=encrypt(catalogItem.iCatalogID)%>" /><input type="hidden" value="<%=true%>" name=postback /><table align=center cellpadding="2"><tr><td colspan=2 class=header><%=l("general")%>:</td></tr><tr><td class=QSlabel><%=l("catalog")%>:</td><td><%=catalogItem.catalog.sName%></td></tr><%if isNumeriek(catalogItem.iCatalogID) then%><tr><td class=QSlabel valign=top><%=catalogItem.catalog.sItemName%>:*</td><td><input type=text size=40 maxlength=250 name="sTitle" value="<%=quotRep(catalogItem.sTitle)%>" /></td></tr><%if isNumeriek(catalogItem.catalog.iFormID) then%><tr><td class=QSlabel valign=top><%=l("useform")%></td><td><input type=checkbox name="bForm" value="1" <%=convertChecked(catalogItem.bForm)%> /></td></tr><%end if%><tr><td class=QSlabel><%=l("date")%>:</td><td><input type="text" id="dDate" name="dDate" value="<%=convertEuroDate(catalogItem.dDate)%>" /><%=JQDatePicker("dDate")%></td></tr><tr><td class=QSlabel><%=l("online")%>&nbsp;<%=l("from")%>:</td><td><input type="text" id="dOnlineFrom" name="dOnlineFrom" value="<%=convertEuroDate(catalogItem.dOnlineFrom)%>" /><%=JQDatePicker("dOnlineFrom")%><i><%=l("untill")%>:</i> 
<input type="text" id="dOnlineUntill" name="dOnlineUntill" value="<%=convertEuroDate(catalogItem.dOnlineUntill)%>" /><%=JQDatePicker("dOnlineUntill")%></td></tr><tr><td colspan=2><hr /></td></tr><%for each catalogField in catalogFields%><tr><td class=QSlabel><%=catalogFields(catalogField).sName%><%if catalogFields(catalogField).bMandatory then%>*<%end if%></td><td><%select case catalogFields(catalogField).sType
case sb_text,sb_url,sb_email%><input type="text" size="50" maxlength="250" name="<%=encrypt(catalogField)%>" value="<%=quotRep(fields(catalogField))%>"><%case sb_textarea%><textarea cols="60" rows="6" name="<%=encrypt(catalogField)%>"><%=quotRep(fields(catalogField))%></textarea><%case sb_checkbox%><input type="checkbox" name="<%=encrypt(catalogField)%>" value="checked" <%=fields(catalogField)%> /><%case sb_select%><select name="<%=encrypt(catalogField)%>"><%=catalogFields(catalogField).showSelected(fields(catalogField))%></select><%case sb_date%><input type="text" id="<%=encrypt(catalogField)%>" name="<%=encrypt(catalogField)%>" value="<%=convertDateToPicker(fields(catalogField))%>" /><%=JQDatePicker(encrypt(catalogField))%><%case sb_richtext%><%createFCKInstance fields(catalogField),"siteBuilderRichText",encrypt(catalogField)%><%end select%></td></tr><%next
if isNumeriek(catalogItem.iID) then%><tr><td colspan=2><hr /></td></tr><tr><td class=QSlabel><%=l("createdon")%>:</td><td><%=convertEuroDateTime(catalogItem.dCreatedTS)%></td></tr><tr><td class=QSlabel><%=l("updatedon")%>:</td><td><%=convertEuroDateTime(catalogItem.dUpdatedTS)%></td></tr><%end if%><tr><td colspan=2><hr /></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=btnaction value="<% =l("save")%>" />

<%if isNumeriek(catalogItem.iID) then%>
<%
if not isLeeg(catalogItem.catalog.sItemView) then
%>
&nbsp;<a target="_blank" class="art-button" href="<%=C_DIRECTORY_QUICKERSITE%>/default.asp?pageAction=showitem&iItemID=<%=encrypt(catalogItem.iId)%>"><% =l("preview")%></a>
<%
end if
%>

&nbsp;<input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>">
<%end if
'if isNumeriek(catalogItem.iID) then
if isLeeg(catalogItem.sPicExt) then%><br /><br /><input class="art-button" type=submit name=btnaction value="<% =l("uploadpic")%>"><%else%><br /><br /><%=catalogItem.showPic("")%><br /><br /><input  class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<%=l("deletepic")%>"><%end if
'end if
if fileTypes.count>0 then ' and isNumeriek(catalogItem.iID) then%><br /><br /><input class="art-button" type=submit name=btnaction value="<% =l("uploadattachment")%>"><%end if%></td></tr><%end if%></table></form><%if files.count>0 then
dim fileKey%><table align=center cellpadding="2"><tr><td colspan=2 class=header><%=l("attachments")%>:</td></tr><%for each filekey in files%><tr><td><%=files(filekey).url%></td><td><a href="#" onclick="javascript:if(confirm('<%=l("areyousure")%>')){location.assign('bs_catalogItemEdit.asp?<%=QS_secCodeURL%>&amp;iItemID=<%=encrypt(catalogItem.iID)%>&iFileID=<%=encrypt(filekey)%>');} else {return false};"><img border=0 alt="<%=l("delete")%>" src="<%=C_DIRECTORY_QUICKERSITE%>/fixedImages/dustbin.gif" /></a></td></tr><%next%></table><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
