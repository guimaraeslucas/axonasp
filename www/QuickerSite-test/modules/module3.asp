<%
' This is a sample form that can be included in any QuickerSite page (using the Application Path variable)
' The most important field here is 'sCode'. It refers to the field 'sCode' of a page in QuickerSite.
' By using the sCode-field, you can make sure the page with the corresponding sCode-value is loaded 
' after your visitor has submitted a form. You can also use the sCode-parameter in a querystring.
%>
<form action="default.asp" method="post">
<input type="hidden" value="FORM" name="sCode">
<table>
	<tr>
		<td>Your name:</td>
		<td><input type="text" name="field" value="<%=server.HTMLEncode (Request.Form ("field"))%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="submit" value="Submit"></td>		
	</tr>
</table>
</form>
<%
if Request.Form ("field")<>"" then
%>
<p>Your name is <%=server.HTMLEncode (Request.Form ("field"))%>.</p>
<%
end if
%>