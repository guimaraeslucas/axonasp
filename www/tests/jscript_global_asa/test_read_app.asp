<%@ Language="JScript" %>
<%
// Reads the value set by global.asa Application_OnStart written in JScript.
// If the fix is working, Application("test") == 1.
var val = Application("test");
if (val !== 1) {
    Response.Write("FAIL: Expected Application(\"test\") = 1, got " + String(val));
} else {
    Response.Write("PASS: Application(\"test\") = " + String(val));
}
%>