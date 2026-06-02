<%@ Language="JScript" %>
<%
var d = new Date(new Date().getTime() + (365 * 24 * 60 * 60 * 1000));
Response.Cookies("test_js_prop") = "prop_value";
Response.Cookies("test_js_prop").Path = "/my_path";
Response.Cookies("test_js_prop").Expires = d;

Response.Write("Cookie set with JScript properties! Date is: " + d);
%>