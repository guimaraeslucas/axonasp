<!-- #include file="begin.asp"-->
<!-- #include file="bs_security.asp"-->
<%
dim fbM

if not convertBool(customer.bListItemPic) then response.end 

dim page
set page=new cls_page
page.pick(decrypt(request("iId")))

if request.form("btnAction")<>"" then
	page.sLPIC=request.form("sLPIC")
	page.save()	
	fbM="<script type=""text/javascript"">window.parent.$.fn.colorbox.close();</script>"
end if

if request.form("delAction")<>"" then
	page.deleteListItemImage()
	fbM="<script type=""text/javascript"">window.parent.$.fn.colorbox.close();</script>"
end if

%>
<html>
<head>
<title>Configure Picture</title>
<%=fbM%>
</head>
<body style="background-color:#FFFFFF">
<form method="post" action="bs_editpictureLI.asp" id=editP name=editP>
<%=QS_secCodeHidden%>
<input  type="hidden" value="<% =EnCrypt(page.iId) %>" name="iId" />
<table style="font-size:11pt" cellspacing="0" cellpadding="3">
	<tr>
		<td><input type="radio" name="sLPIC" <%if page.sLPIC="al" then response.write " checked=""checked"" "%> value="al" /></td><td>Align left (50%)</td>
	</tr>
	<tr>
		<td><input type="radio" name="sLPIC" <%if page.sLPIC="fp" then response.write " checked=""checked"" "%> <%if isLeeg(page.sLPIC) then response.write " checked=""checked"" "%> value="fp" /></td><td>Full page width (100%)</td>
	</tr>
	<tr>
		<td><input type="radio" name="sLPIC" <%if page.sLPIC="ar" then response.write " checked=""checked"" "%> value="ar" /></td><td>Align right (50%)</td>
	</tr>
	<tr>
		<td><input type="radio" name="sLPIC" <%if page.sLPIC="CC" then response.write " checked=""checked"" "%> value="CC" /></td><td>custom CSS on class="ListItemPictureCSS"</td>
	</tr>
	<tr>
		<td colspan="2"><input type=submit value="Save" name="btnAction" />  <input type=submit onclick="javascript:return confirm('Are you sure to delete the picture?');" value="Delete Picture" name="delAction" /></td>
	</tr>
</table>
</form>
</body>
</html>
<%cleanUPASP%>
