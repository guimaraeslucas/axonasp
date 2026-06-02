<%@ Language="JScript" %>
<%
var c = Request.Cookies("existing_cookie");
var str = c + "";
Response.Write("Value: " + str);
%>