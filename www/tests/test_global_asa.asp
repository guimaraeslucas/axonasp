<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head>
    <title>Global.asa Test - G3pix AxonASP</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #333; }
        h2 { color: #666; margin-top: 20px; }
        .info { background: #f0f0f0; padding: 10px; margin: 10px 0; border-left: 4px solid #4CAF50; }
        .success { color: #4CAF50; }
        .code { background: #272822; color: #f8f8f2; padding: 10px; font-family: monospace; margin: 10px 0; }
    </style>
</head>
<body>
    <h1>Global.asa Event Test</h1>
    
    <div class="info">
        <p><strong>This test verifies that global.asa is properly loaded and executed.</strong></p>
        <p>The global.asa file contains Application_OnStart and Session_OnStart event handlers.</p>
    </div>
    
    <h2>Application State</h2>
    <div class="info">
        <p><strong>Application("visitors"):</strong> <span class="success"><%= Application("visitors") %></span></p>
        <p>This value is initialized to 0 in Application_OnStart and incremented in Session_OnStart.</p>
    </div>
    
    <h2>Session Information</h2>
    <div class="info">
        <p><strong>Session ID:</strong> <%= Session.SessionID %></p>
        <p><strong>Session Timeout:</strong> <%= Session.Timeout %> minutes</p>
    </div>
    
    <h2>Global.asa Events</h2>
    <div class="info">
        <p><strong>✓ Application_OnStart:</strong> Executed once when the server starts</p>
        <p><strong>✓ Session_OnStart:</strong> Executed when a new session is created</p>
        <p><strong>✓ Session_OnEnd:</strong> Executed when a session expires or is abandoned</p>
        <p><strong>✓ Application_OnEnd:</strong> Executed when the server stops</p>
    </div>
    
    <h2>Global.asa Code</h2>
    <div class="code">
&lt;script language="vbscript" runat="server"&gt;<br>
<br>
Sub Application_OnStart<br>
&nbsp;&nbsp;&nbsp;&nbsp;Application("visitors")=0<br>
End Sub<br>
<br>
Sub Session_OnStart<br>
&nbsp;&nbsp;&nbsp;&nbsp;Application.Lock<br>
&nbsp;&nbsp;&nbsp;&nbsp;Application("visitors")=Application("visitors")+1<br>
&nbsp;&nbsp;&nbsp;&nbsp;Application.UnLock<br>
End Sub<br>
<br>
Sub Session_OnEnd<br>
&nbsp;&nbsp;&nbsp;&nbsp;Application.Lock<br>
&nbsp;&nbsp;&nbsp;&nbsp;Application("visitors")=Application("visitors")-1<br>
&nbsp;&nbsp;&nbsp;&nbsp;Application.UnLock<br>
End Sub<br>
<br>
&lt;/script&gt;
    </div>
    
    <hr>
    <p><a href="test_global_asa.asp">Refresh Page (New Session)</a> | <a href="../default.asp">Back to Home</a></p>
    
    <p style="color: #999; font-size: 12px; margin-top: 30px;">
        <em>Note: Each refresh without cookies creates a new session, incrementing the visitor count.</em>
    </p>
</body>
</html>
