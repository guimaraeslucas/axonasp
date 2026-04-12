<%
@ Language = "VBScript"
%>
<h2>Simple Session Test</h2>
<p>Session_OnStart values (if set):</p>
<ul>
    <li>
        Time:
        <%= Session("Global_SessionStart_Time") %>
    </li>
    <li>
        Message:
        <%= Session("Global_SessionStart_Msg") %>
    </li>
    <li>
        Session ID:
        <%= Session.SessionID %>
    </li>
</ul>
