<%@ Language="JScript" %>
<%
var q = Request.QueryString("lang");
var str = q + "";
Response.Write("Value: " + str);
%>