<!-- #include file="begin.asp"-->
<!-- #include file="bs_security.asp"-->
<!-- #include file="includes/header.asp"-->
<!-- #include file="bs_initBack.asp"-->
<!-- #include file="bs_header.asp"-->
<%=getBOHeader("")%>
<%
dim page
set page=new cls_page
page.pick(decrypt(request("iId")))

if request.form("sLPTemplate")<>"" then
page.sLPTemplate=request.form("sLPTemplate")
if page.save() then response.redirect ("bs_listPage.asp?iID=" & encrypt(page.iID))
end if
%>

<table cellpadding="10" align=center><tr><td align=center>-> <b><a href="bs_listPage.asp?iID=<%=encrypt(page.iID)%>"><%=l("back")%></a></b> <-</td></tr></table>

<form action="bs_editListTemplate.asp" method="post" name=mainform>
<%=QS_secCodeHidden%>
<INPUT type="hidden" value="<% =EnCrypt(page.iId) %>" name=iId />
<table align=center style="height:80%" width=640 cellpadding="2">
<tr><td colspan=2 class=header>Select display for list page "<%=page.sTitle%>":</td></tr>

<tr>
	<td style="text-align:center"><img style="margin:10px;width:200px;box-shadow:1px 1px 3px #222" src="<%=C_DIRECTORY_QUICKERSITE%>/fixedImages/liststyles/ls1.jpg" />
	<input <%if page.sLPTemplate="1" then response.write " checked=""checked"" "%> type="radio" name="sLPTemplate" value="1" />
	</td>
	<td style="text-align:center"><img style="margin:10px;width:200px;box-shadow:1px 1px 3px #222" src="<%=C_DIRECTORY_QUICKERSITE%>/fixedImages/liststyles/ls2.jpg" />
	<input <%if page.sLPTemplate="2" then response.write " checked=""checked"" "%> type="radio" name="sLPTemplate" value="2" />
	</td>
	<td style="text-align:center"><img style="margin:10px;width:200px;box-shadow:1px 1px 3px #222" src="<%=C_DIRECTORY_QUICKERSITE%>/fixedImages/liststyles/ls3.jpg" />
	<input <%if page.sLPTemplate="3" then response.write " checked=""checked"" "%> type="radio" name="sLPTemplate" value="3" />
	</td>
</tr>


<tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr>
<tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=dummy value="<% =l("save")%>" />&nbsp;<input class="art-button" type=reset  value="<% =l("reset")%>" id=reset1 name=reset1 />
</td></tr>
</table>
</form>

<table cellpadding="10" align=center><tr><td align=center>-> <b><a href="bs_listPage.asp?iID=<%=encrypt(page.iID)%>"><%=l("back")%></a></b> <-</td></tr></table>

<!-- #include file="bs_endBack.asp"-->
<!-- #include file="includes/footer.asp"-->
