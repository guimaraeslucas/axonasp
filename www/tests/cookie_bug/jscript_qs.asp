<%@ Language="JScript" %>
<%
var q = Request.QueryString("lang");
var str = String(q);
Response.Write("String(q) = '" + str + "'");
%>