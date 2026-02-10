<!-- #include file="begin.asp"-->
<%pagetoemail=true
editPage=true
dim getLperP
set getLperP=logon.contact.getLper

if not getLperP.exists(selectedPage.iId) then Response.Redirect (C_DIRECTORY_QUICKERSITE & "/asp/noaccess.htm")%>

<!-- #include file="includes/commonheader.asp"-->

<body style="color:#000;background-color:#FFF;background-image:none" onload="javascript:window.focus();">
<%
dim listitems
set listitems=selectedPage.listitems(false)
%>
<table align=center  cellpadding="5" style="background-color:#FFF">
<tr><td><%=l("managelist")%>: '<b><%=selectedPage.sTitle%></b>'</td><td>
<form action="fs_editListItem.asp" method="post">
<input type="hidden" name="iListPageID" value="<%=encrypt(selectedPage.iID)%>" />
<input type="submit" class="art-button" value="<%=l("addnewitem")%>" />
</form></tr>
<tr><td colspan="2">
<ul>
<li style="margin:6px"><%=l("modifylistitem")%><%if listitems.count>0 then%>
<ul>
<%dim item
for each item in listitems

if listitems(item).statusString=l("offline") then
Response.Write "<li style=""margin:6px""><i>Offline:&nbsp;<a href=""fs_editListItem.asp?iId="& encrypt(item) & """>" & listitems(item).sTitle & "</i></a>"
else
Response.Write "<li style=""margin:6px""><a href=""fs_editListItem.asp?iId="& encrypt(item) & """>" & listitems(item).sTitle & "</a>"
end if

if listitems(item).statusString<>l("offline") then
Response.Write "&nbsp;-&nbsp;<a target=""_blank"" href=""" & C_DIRECTORY_QUICKERSITE & "/default.asp?iId=" & encrypt(listitems(item).iListPageID) & "&amp;item=" & encrypt(item) & "#" & encrypt(item) & """>" & l("preview") & "</a>"
end if

if convertBool(customer.bListItemPic) then
response.write "&nbsp;-&nbsp;<a href=""fs_editListItemPic.asp?iId=" & encrypt(item) & """>Set/Edit picture</a>"
end if

response.write "</li>"

next%>
</ul></li>
<%end if%>
</ul></td></tr>
</table></body></html>
<%cleanUPASP%>
