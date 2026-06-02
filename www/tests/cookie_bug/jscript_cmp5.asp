<%@ Language="JScript" %>
<%
var c = Request.Cookies("existing_cookie");
var result = "";
if (c == "random_garbage") {
    result += "c == 'random_garbage' is true. ";
} else {
    result += "c == 'random_garbage' is false. ";
}
Response.Write(result);
%>