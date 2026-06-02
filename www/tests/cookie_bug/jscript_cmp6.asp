<%@ Language="JScript" %>
<%
var c = Request.Cookies("existing_cookie");
var result = "";
if (c == "") result += "c=='' is true. ";
if (c == "my_value") result += "c=='my_value' is true. ";
Response.Write(result);
%>