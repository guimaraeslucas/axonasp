<!-- #include file="begin.asp"-->
<!-- #include file="bs_security.asp"-->
<!-- #include file="bs_process.asp"-->
<!-- #include file="includes/header.asp"-->
<!-- #include file="bs_initBack.asp"-->
<!-- #include file="bs_header.asp"--><%=getBOHeader("")%>
<%
if convertGetal(decrypt(request.querystring("iDeleteLIID")))<>0 then
	checkCSRf()
	set listitem=new cls_page
	listitem.pick(convertGetal(decrypt(request.querystring("iDeleteLIID"))))
	if convertGetal(listitem.iId)<>0 then
		listitem.bDeleted=true
		listitem.bOnline=false
		if listitem.save then
			Response.Redirect ("bs_listPage.asp?fbMessage=fb_topicremoved&iID=" & encrypt(listitem.iListPageID))
		end if
	end if
end if

dim listitem
if convertGetal(decrypt(request.querystring("iCopyIid")))<>0 then
	checkCSRf()
	
	set listitem=new cls_page
	listitem.pick(convertGetal(decrypt(request.querystring("iCopyIid"))))
	if convertGetal(listitem.iId)<>0 then
		listitem.iId=null
		listitem.sCode=""
		listitem.sUserFriendlyURL=""
		listitem.sTitle=listitem.sTitle & " (copy)"
		if listitem.save then
			Response.Redirect ("bs_listPage.asp?iID=" & encrypt(listitem.iListPageID))
		end if
	end if
end if

set listitem=new cls_page


dim listitems, item
set listitems=page.listitems(false)%><form name="mainform" method="post" action="bs_listPage.asp"><%=QS_secCodeHidden%><table align=center cellpadding="2"><tr><td><%=l("managelist")%>: '<b><%=page.sTitle%></b>'
<p><a class="art-button" href="bs_editList.asp?iID=<%=encrypt(page.iID)%>"><%=l("expllistpage")%></a>&nbsp;
<a class="art-button" href="bs_editListItem.asp?iListPageID=<%=encrypt(page.iID)%>"><%=l("addnewitem")%></a></p><p><%=l("modifylistitem")%></p><%if listitems.count>0 then%><ul>
<%for each item in listitems
Response.Write "<li style=""margin-bottom:10px"">" & listitems(item).getClickLink(true) & " <i>- iId: " & item & "</i>"
Response.Write "&nbsp;<a class=""art-button"" href=""bs_editListItem.asp?iId="& encrypt(item)& """>" & l("modify") & "</a>"

if not customer.bUserFriendlyURL or isLeeg(listitems(item).sUserfriendlyURL) then 
	Response.Write "&nbsp;<a target=""_blank"" class=""art-button"" href=""" & C_DIRECTORY_QUICKERSITE & "/default.asp?iId=" & encrypt(listitems(item).iListPageID) & "&amp;item=" & encrypt(item) & "#" & encrypt(item) & """>" & l("preview") & "</a>"
else
	Response.Write "&nbsp;<a target=""_blank"" class=""art-button"" href=""" & C_VIRT_DIR & "/" & listitems(item).sUserfriendlyURL & """>" & l("preview") & "</a>"	
end if

'this is for the pictures in the list items
if convertBool(customer.bListItemPic) then
	if not isLeeg(listitems(item).sItemPicture) then
		Response.Write "&nbsp;<a class=""art-button QSPPIMG"" href=""" & C_VIRT_DIR &  Application("QS_CMS_userfiles") & "listitemimages/" & item & "." & listitems(item).sItemPicture & """>View picture</a>"
		Response.Write "&nbsp;<a onclick=""$.colorbox({height:'240', width:'420', iframe:true, "
		Response.Write "href:'bs_editpictureLI.asp?iId=" & encrypt(item) & "', "
		Response.Write "onClosed:function(){location.reload(true);}});"" class=""art-button"" href=""#"">Configure picture</a>"
	else
		Response.Write "&nbsp;<a onclick=""$.colorbox({height:'240', width:'420', iframe:true, "
		Response.Write "href:'bs_uploadpictureLI.asp?iId=" & encrypt(item) & "', "
		Response.Write "onClosed:function(){location.reload(true);}});"" class=""art-button"" href=""#"">" & l("picture") & "</a>"
	end if
end if

Response.Write "&nbsp;<a onclick=""javascript:return confirm('" & l("areyousure") & "');"" "
Response.Write "href=""bs_ListPage.asp?QSSEC=" & secCode & "&amp;iId=" & encrypt(page.iId) & "&amp;iCopyIID=" 
Response.Write encrypt(item) & """ class=""art-button"">" & l("copy") & "</a>"

Response.Write "&nbsp;<a onclick=""javascript:return confirm('" & l("deletecomplete") & "');"" "
Response.Write "href=""bs_ListPage.asp?QSSEC=" & secCode & "&amp;iId=" & encrypt(page.iId) & "&amp;iDeleteLIID=" 
Response.Write encrypt(item) & """ class=""art-button"">" & l("delete") & "</a>"


Response.Write "</li>"
next%></ul><%end if%></ul></td></tr></table></form><!-- #include file="bs_editPageBlocksInc.asp"--><p align=center>-> <b><a onclick="javascript:return confirm('<%=l("areyousure")%>');" href="bs_listPage.asp?<%=QS_secCodeURL%>&amp;btnaction=makenormalpage&amp;iID=<%=encrypt(page.iID)%>"><%=l("makenormalpage")%></a></b> <-</p><p align=center>-> <b><a target="<%=encrypt(page.iId)%>" href="<%=C_DIRECTORY_QUICKERSITE%>/default.asp?iID=<%=encrypt(page.iID)%>"><%=l("check")%>&nbsp;<%=page.sTitle%></a></b> <-</p>
<!-- #include file="bs_back.asp"-->
<!-- #include file="bs_endBack.asp"-->
<!-- #include file="includes/footer.asp"-->