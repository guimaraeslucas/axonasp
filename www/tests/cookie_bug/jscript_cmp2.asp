<%@ Language="JScript" %>
<%
var c = Request.Cookies("existing_cookie");
Response.Write("String(c) = '" + String(c) + "'");
%>