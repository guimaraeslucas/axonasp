<%@ Language="JScript" %>
<%
var c = Request.Cookies("existing_cookie");
var result = "";
if (c == "my_value") {
    result += "c == 'my_value' is true. ";
} else {
    result += "c == 'my_value' is false. ";
}
Response.Write(result);
%>