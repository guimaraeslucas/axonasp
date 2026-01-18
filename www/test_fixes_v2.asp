<%@ Language=VBScript %>
<%
Response.Write "<h3>Test Fixes V2</h3>"
%>
<p>Testing Include Virtual (footer.inc):</p>
<div style="border:1px solid #ccc; padding:10px;">
<!--#include virtual="/footer.inc"-->
</div>

<p>Testing Include File (relative header.inc):</p>
<div style="border:1px solid #ccc; padding:10px;">
<!--#include file="header.inc"-->
</div>

<p>Testing Session Persistence:</p>
<%
Dim count
count = Session("FixTestCounter")
If IsEmpty(count) Or count = "" Then
    count = 1
    Session("FixTestCounter") = 1
    Response.Write "First visit (Counter initialized to 1). Refresh to increment."
Else
    count = CInt(count) + 1
    Session("FixTestCounter") = count
    Response.Write "Counter: " & count
End If
%>

<p>Testing Redirect (Click link to test):</p>
<a href="test_redirect_target.asp">Go to Redirect Test</a>
