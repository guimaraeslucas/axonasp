<%@ Language="JScript" %>
<%
function SafeToString(value) {
    if (value == null) return "";
    return String(value);
}
var s = Session("g3pix_lang");
var res = SafeToString(s);
Response.Write("Session is: [" + res + "]");
%>