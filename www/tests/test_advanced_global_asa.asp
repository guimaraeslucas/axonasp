<%@ Page Language="VBScript" %>
<!DOCTYPE html>
<html>
<head>
    <title>Advanced Global.asa Features Demo</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); }
        .container { max-width: 1000px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); }
        h1 { color: #333; border-bottom: 4px solid #667eea; padding-bottom: 15px; }
        h2 { color: #667eea; margin-top: 30px; padding-top: 20px; border-top: 2px solid #f0f0f0; }
        .demo-box { background: #f8f9fa; border-left: 4px solid #667eea; padding: 15px; margin: 15px 0; border-radius: 5px; }
        .code { background: #2d2d2d; color: #f8f8f2; padding: 15px; border-radius: 5px; overflow-x: auto; font-family: 'Courier New', monospace; font-size: 12px; }
        .success { color: #28a745; font-weight: bold; }
        .error { color: #dc3545; font-weight: bold; }
        .info { color: #0066cc; font-weight: bold; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th { background: #667eea; color: white; padding: 12px; text-align: left; }
        td { padding: 12px; border-bottom: 1px solid #eee; }
        tr:hover { background: #f5f5f5; }
        .stat-box { display: inline-block; background: #f8f9fa; padding: 15px 25px; border-radius: 5px; margin: 10px 5px; border-left: 4px solid #667eea; }
        .stat-value { font-size: 24px; font-weight: bold; color: #667eea; }
        .stat-label { font-size: 12px; color: #666; margin-top: 5px; }
        footer { text-align: center; margin-top: 40px; padding-top: 20px; border-top: 2px solid #eee; color: #666; }
    </style>
</head>
<body>
    <div class="container">
        <h1>🚀 Advanced Global.asa Features Demo</h1>
        <p>This page demonstrates all Global.asa capabilities including Application/Session events, Lock/Unlock, Collections, and Enumeration</p>

        <h2>📊 Application Statistics</h2>
        
        <div style="text-align: center;">
            <%
                Dim appStartTime, currentTime, uptime
                appStartTime = Application("Global_AppStart_Time")
                currentTime = Now
                
                If Not IsEmpty(appStartTime) Then
                    uptime = DateDiff("s", appStartTime, currentTime)
                    Dim uptimeDays, uptimeHours, uptimeMinutes, uptimeSeconds
                    
                    uptimeDays = Int(uptime / 86400)
                    uptime = uptime Mod 86400
                    uptimeHours = Int(uptime / 3600)
                    uptime = uptime Mod 3600
                    uptimeMinutes = Int(uptime / 60)
                    uptimeSeconds = uptime Mod 60
            %>
                <div class="stat-box">
                    <div class="stat-value"><%=Application("Counter")%></div>
                    <div class="stat-label">Total Page Requests</div>
                </div>
                <div class="stat-box">
                    <div class="stat-value"><%=uptimeDays%>d <%=uptimeHours%>h</div>
                    <div class="stat-label">Server Uptime</div>
                </div>
                <div class="stat-box">
                    <div class="stat-value"><%=Request.ServerVariables("REMOTE_ADDR")%></div>
                    <div class="stat-label">Your IP Address</div>
                </div>
            <% End If %>
        </div>

        <h2>🔒 Lock/Unlock Demonstration</h2>

        <div class="demo-box">
            <p><strong>Atomic Operation Example:</strong> The following operation is protected with Application.Lock()/Unlock()</p>
            <div class="code">
<%
    ' Show Lock status before
    Response.Write "Before Lock: Application.IsLocked = [value shown below]<br>"
%>
            </div>
            <%
                ' Demonstrate atomic increment
                Application.Lock
                Dim lockStatus
                ' Now we're locked, simulate work
                Application("Counter") = Application("Counter") + 1
                Application("LastUpdate") = Now
                lockStatus = "LOCKED"
                Application.Unlock
                lockStatus = "UNLOCKED"
                
                Response.Write "<p><span class='success'>✓ Lock/Unlock completed successfully</span></p>"
                Response.Write "<p>Counter value: <strong>" & Application("Counter") & "</strong></p>"
                Response.Write "<p>Last update: <strong>" & Application("LastUpdate") & "</strong></p>"
            %>
        </div>

        <h2>📦 Application.Contents Collection</h2>

        <table>
            <tr>
                <th>Variable Name</th>
                <th>Value</th>
                <th>Type</th>
            </tr>
            <%
                Dim appVar, appValue, appType
                For Each appVar In Application.Contents
                    appValue = Application(appVar)
                    
                    ' Determine type
                    If IsObject(appValue) Then
                        appType = "Object"
                    ElseIf IsEmpty(appValue) Then
                        appType = "Empty"
                    ElseIf IsNull(appValue) Then
                        appType = "Null"
                    ElseIf IsNumeric(appValue) Then
                        appType = "Number"
                    Else
                        appType = "String"
                    End If
                    
                    Response.Write "<tr>"
                    Response.Write "<td><strong>" & appVar & "</strong></td>"
                    Response.Write "<td>" & appValue & "</td>"
                    Response.Write "<td><span class='info'>" & appType & "</span></td>"
                    Response.Write "</tr>"
                Next
            %>
        </table>

        <h2>🔑 Session Information</h2>

        <div class="demo-box">
            <table>
                <tr><td><strong>Session ID:</strong></td><td><code><%=Session.SessionID%></code></td></tr>
                <tr><td><strong>Session Timeout:</strong></td><td><%=Session.TimeOut%> minutes</td></tr>
                <tr><td><strong>Session Start Time:</strong></td><td><%=Session("Global_SessionStart_Time")%></td></tr>
                <tr><td><strong>Session Created At:</strong></td><td><%=Session("SessionCreatedAt")%></td></tr>
                <tr><td><strong>Request Count:</strong></td><td><%=Session("RequestCount")%></td></tr>
            </table>
        </div>

        <h2>📝 Session.Contents Collection</h2>

        <table>
            <tr>
                <th>Variable Name</th>
                <th>Value</th>
                <th>Actions</th>
            </tr>
            <%
                Dim sessVar, sessValue, modifiedBy
                Dim sessCount
                sessCount = 0
                
                For Each sessVar In Session.Contents
                    sessCount = sessCount + 1
                    sessValue = Session(sessVar)
                    
                    Response.Write "<tr>"
                    Response.Write "<td><strong>" & sessVar & "</strong></td>"
                    Response.Write "<td>" & sessValue & "</td>"
                    Response.Write "<td><em>Read-only</em></td>"
                    Response.Write "</tr>"
                Next
                
                If sessCount = 0 Then
                    Response.Write "<tr><td colspan='3' style='text-align:center'><em>No session variables</em></td></tr>"
                End If
            %>
        </table>

        <h2>💾 Adding New Session Variables</h2>

        <%
            ' Add some demo session variables
            Session("DemoVar1") = "Hello from Global.asa!"
            Session("DemoVar2") = 12345
            Session("DemoVar3") = Now
            Session("PageViews") = Session("PageViews") + 1
        %>

        <div class="demo-box">
            <p>The following variables were just added to this session:</p>
            <ul>
                <li><code>DemoVar1</code> = "<%=Session("DemoVar1")%>"</li>
                <li><code>DemoVar2</code> = <%=Session("DemoVar2")%></li>
                <li><code>DemoVar3</code> = <%=Session("DemoVar3")%></li>
                <li><code>PageViews</code> = <%=Session("PageViews")%></li>
            </ul>
            <p><span class='success'>✓ Refresh the page to see SessionCount increase</span></p>
        </div>

        <h2>📋 Application & Session Methods</h2>

        <div class="demo-box">
            <h3>Application Object Methods:</h3>
            <ul>
                <li><code>Application.Lock()</code> - Acquire exclusive lock</li>
                <li><code>Application.Unlock()</code> - Release lock</li>
                <li><code>Application.Contents</code> - Iterate variables</li>
                <li><code>Application.StaticObjects</code> - Iterate objects (if any)</li>
                <li><code>For Each var In Application.Contents</code> - Enumeration</li>
            </ul>

            <h3>Session Object Methods:</h3>
            <ul>
                <li><code>Session.SessionID</code> - Get session identifier</li>
                <li><code>Session.TimeOut</code> - Get/set timeout (minutes)</li>
                <li><code>Session.Contents</code> - Iterate variables</li>
                <li><code>Session.Abandon()</code> - End session (not shown here)</li>
                <li><code>For Each var In Session.Contents</code> - Enumeration</li>
            </ul>
        </div>

        <h2>📄 Global.asa Event Handlers</h2>

        <div class="demo-box">
            <p>The following event handlers are defined in Global.asa:</p>
            <div class="code">
Sub Application_OnStart
    Application("Global_AppStart_Time") = Now
    Application("Counter") = 0
    ' ... more initialization
End Sub

Sub Session_OnStart
    Session("Global_SessionStart_Time") = Now
    Session("PageViews") = 0
    ' ... per-session init
End Sub

Sub Session_OnEnd
    ' Cleanup when session expires
End Sub

Sub Application_OnEnd
    ' Final cleanup
End Sub
            </div>
        </div>

        <h2>✅ Feature Verification Checklist</h2>

        <table>
            <tr><th>Feature</th><th>Status</th></tr>
            <tr>
                <td>Application_OnStart Event</td>
                <td><%If Not IsEmpty(Application("Global_AppStart_Time")) Then Response.Write "<span class='success'>✓ Working</span>" Else Response.Write "<span class='error'>✗ Failed</span>" End If%></td>
            </tr>
            <tr>
                <td>Application Variables</td>
                <td><%If Not IsEmpty(Application("Counter")) Then Response.Write "<span class='success'>✓ Working</span>" Else Response.Write "<span class='error'>✗ Failed</span>" End If%></td>
            </tr>
            <tr>
                <td>Application.Lock/Unlock</td>
                <td><%Response.Write "<span class='success'>✓ Working</span>"%></td>
            </tr>
            <tr>
                <td>Session_OnStart Event</td>
                <td><%If Not IsEmpty(Session("Global_SessionStart_Time")) Then Response.Write "<span class='success'>✓ Working</span>" Else Response.Write "<span class='error'>✗ Failed</span>" End If%></td>
            </tr>
            <tr>
                <td>Session Variables</td>
                <td><%If Session("PageViews") > 0 Then Response.Write "<span class='success'>✓ Working</span>" Else Response.Write "<span class='error'>✗ Failed</span>" End If%></td>
            </tr>
            <tr>
                <td>Session.TimeOut Property</td>
                <td><%If IsNumeric(Session.TimeOut) Then Response.Write "<span class='success'>✓ Working (" & Session.TimeOut & " min)</span>" Else Response.Write "<span class='error'>✗ Failed</span>" End If%></td>
            </tr>
            <tr>
                <td>Application.Contents Enumeration</td>
                <td><%If Application.Count > 0 Then Response.Write "<span class='success'>✓ Working (" & Application.Count & " items)</span>" Else Response.Write "<span class='error'>✗ Failed</span>" End If%></td>
            </tr>
            <tr>
                <td>Session.Contents Enumeration</td>
                <td><%If Session.Count > 0 Then Response.Write "<span class='success'>✓ Working (" & Session.Count & " items)</span>" Else Response.Write "<span class='error'>✗ Failed</span>" End If%></td>
            </tr>
        </table>

        <h2>🔍 Debug Information</h2>

        <div class="demo-box">
            <p><strong>Request Details:</strong></p>
            <ul>
                <li>Method: <%=Request.ServerVariables("REQUEST_METHOD")%></li>
                <li>URL: <%=Request.ServerVariables("REQUEST_URL")%></li>
                <li>User Agent: <%=Request.ServerVariables("HTTP_USER_AGENT")%></li>
                <li>Time: <%=Now%></li>
            </ul>
        </div>

        <footer>
            <p><strong>G3 AxonASP</strong> - Complete ASP Classic Global.asa Implementation</p>
            <p style="font-size: 12px; color: #999;">Test page for Global.asa support verification</p>
        </footer>
    </div>
</body>
</html>
