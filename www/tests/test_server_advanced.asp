<%@ Language="VBScript" %>
<!--#include file="header.inc"-->

<h2>Server Object Advanced Test</h2>

<%
    Dim ParentVar
    ParentVar = "I am defined in the parent"
    
    ' Test ScriptTimeout
    Server.ScriptTimeout = 45
%>

<div class="test-box">
    <h3>Properties</h3>
    <p>Server.ScriptTimeout set to: <b><%= Server.ScriptTimeout %></b> seconds.</p>
</div>

<div class="test-box">
    <h3>Server.Execute Test</h3>
    <p>This method executes a child page and returns control to this page.</p>
    
    <hr>
    <b>[Before Execute]</b>
    <% Server.Execute "test_server_child.asp" %>
    <b>[After Execute]</b> <span style="color: green;">(If you see this, control returned successfully)</span>
    <hr>
</div>

<div class="test-box">
    <h3>Server.Transfer Test</h3>
    <p><a href="test_server_advanced.asp?mode=transfer" class="button">Click to Test Server.Transfer</a></p>
    
    <% If Request.QueryString("mode") = "transfer" Then %>
        <hr>
        <b>[Before Transfer]</b>
        <% 
            Server.Transfer "test_server_child.asp" 
            ' Code below this should NOT run
        %>
        <br>
        <b style="color: red;">[FAIL] This line should NOT be visible!</b>
    <% End If %>
</div>

<!--#include file="footer.inc"-->
