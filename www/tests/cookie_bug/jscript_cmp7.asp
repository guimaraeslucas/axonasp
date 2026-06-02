<%@ Language="JScript" %>
<%
var c = Request.Cookies("existing_cookie");
var result = "";
try {
    var str = String(c);
    result += "String(c) worked: " + str;
} catch (e) {
    result += "String(c) threw: " + e.message;
}
Response.Write(result);
%>