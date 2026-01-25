<%@ Language="VBScript" %>
<!--#include file="header.inc"-->

<h2>Global.asa Event Test</h2>

<div class="test-box">
    <h3>Application_OnStart</h3>
    <p>If working, you should see a timestamp and message below (requires server restart if this is the first run):</p>
    <ul>
        <li><b>Time:</b> <%= Application("Global_AppStart_Time") %></li>
        <li><b>Message:</b> <%= Application("Global_AppStart_Msg") %></li>
    </ul>
</div>

<div class="test-box">
    <h3>Session_OnStart</h3>
    <p>If working, you should see a timestamp and message below (triggers on new session):</p>
    <ul>
        <li><b>Time:</b> <%= Session("Global_SessionStart_Time") %></li>
        <li><b>Message:</b> <%= Session("Global_SessionStart_Msg") %></li>
        <li><b>Session ID:</b> <%= Session.SessionID %></li>
    </ul>
    
    <p>
        <a href="test_global_events.asp" class="button">Refresh Page</a>
        <a href="test_global_events.asp?action=abandon" class="button">Abandon Session & Refresh</a>
    </p>

    <%
    If Request.QueryString("action") = "abandon" Then
        Session.Abandon
        Response.Write "<p><i>Session Abandoned. Click refresh to see new session start.</i></p>"
    End If
    %>
</div>

<!--#include file="footer.inc"-->
