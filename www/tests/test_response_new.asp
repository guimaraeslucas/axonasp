<%@ Language="VBScript" %>
<%
    ' Test Response Properties
    Response.Expires = 5
    Response.ExpiresAbsolute = Now
    Response.CacheControl = "Public"
    Response.Charset = "ISO-8859-1"
    
    ' Test Response Methods
    Response.AddHeader "X-Custom-Test", "AxonASP-Value"
    Response.AppendToLog "User visited test_response_new.asp at " & Now
%>
<!--#include file="header.inc"-->

<h2>Response Object Upgrades Test</h2>

<div class="test-box">
    <h3>Headers & Properties</h3>
    <p>We have set the following server-side properties. Open your browser's <b>DevTools (Network Tab)</b> to verify response headers:</p>
    <ul>
        <li><code>Response.Expires = 5</code></li>
        <li><code>Response.CacheControl = "Public"</code></li>
        <li><code>Response.Charset = "ISO-8859-1"</code> (Check Content-Type header)</li>
        <li><code>Response.AddHeader "X-Custom-Test", ...</code></li>
    </ul>
</div>

<div class="test-box">
    <h3>BinaryWrite Test</h3>
    <p>Below is content written via <code>Response.BinaryWrite</code>:</p>
    <div style="border: 1px solid #ccc; padding: 10px; background: #fff;">
    <%
        Response.BinaryWrite "<b>This text was written using BinaryWrite.</b>"
    %>
    </div>
</div>

<div class="test-box">
    <h3>Logging</h3>
    <p><code>Response.AppendToLog</code> was called. Check the server console output (stdout) for the log message.</p>
</div>

<!--#include file="footer.inc"-->
