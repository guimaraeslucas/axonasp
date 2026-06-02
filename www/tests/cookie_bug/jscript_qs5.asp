<%@ Language="JScript" %>
<%
function TrimString(value) {
    if (value == null) return "WAS_NULL";
    var str = String(value);
    return "STR_IS: [" + str + "]";
}
var c = Request.Cookies("existing_cookie");
var res = TrimString(c);
Response.Write(res);
%>