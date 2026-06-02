<%@ Language="JScript" %>
<%
function TrimString(value) {
    if (value == null) return "WAS_NULL";
    var str = String(value);
    return "STR_IS: [" + str + "]";
}
var q = Request.QueryString("lang");
var res = TrimString(q);
Response.Write(res);
%>