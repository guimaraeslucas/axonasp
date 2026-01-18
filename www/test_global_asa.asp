<%@ Page Language="VBScript" %>
<!DOCTYPE html>
<html>
<head>
    <title>Global.asa Support Test</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; }
        h1 { color: #0066cc; border-bottom: 3px solid #0066cc; padding-bottom: 10px; }
        h2 { color: #333; margin-top: 30px; border-left: 4px solid #0066cc; padding-left: 10px; }
        .section { margin: 20px 0; padding: 15px; background: #f9f9f9; border-left: 4px solid #0066cc; }
        .success { color: #28a745; font-weight: bold; }
        .error { color: #dc3545; font-weight: bold; }
        .info { color: #0066cc; font-weight: bold; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; }
        table td { padding: 10px; border: 1px solid #ddd; }
        table td:first-child { font-weight: bold; background: #f0f0f0; width: 350px; }
        .value { color: #666; font-family: monospace; background: #f0f0f0; padding: 4px 8px; border-radius: 4px; }
        code { background: #f0f0f0; padding: 2px 6px; border-radius: 3px; font-family: monospace; }
        pre { background: #f0f0f0; padding: 10px; border-radius: 4px; overflow-x: auto; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Global.asa Support Test</h1>
        <p>Testing ASP Classic Global.asa functionality including Application_OnStart, Session_OnStart, Lock/Unlock, and StaticObjects</p>

        <h2>Application Object Tests</h2>

        <div class="section">
            <h3>Application Variables from Global.asa</h3>
            <table>
                <tr><td>Global_AppStart_Time</td><td>
                    <%
                        Dim appStartTime
                        Set appStartTime = Application("Global_AppStart_Time")
                        If Not IsEmpty(appStartTime) And Not IsNull(appStartTime) Then
                            Response.Write "<span class='success'>Set: " & appStartTime & "</span>"
                        Else
                            Response.Write "<span class='error'>NOT SET</span>"
                        End If
                    %>
                </td></tr>
                <tr><td>Global_AppStart_Msg</td><td>
                    <%
                        Dim appStartMsg
                        Set appStartMsg = Application("Global_AppStart_Msg")
                        If Not IsEmpty(appStartMsg) And Not IsNull(appStartMsg) Then
                            Response.Write "<span class='success'>" & Application("Global_AppStart_Msg") & "</span>"
                        Else
                            Response.Write "<span class='error'>NOT SET</span>"
                        End If
                    %>
                </td></tr>
            </table>
        </div>

        <div class="section">
            <h3>Application.Lock() / Unlock() Test</h3>
            <table>
                <tr><td>Calling Application.Lock()</td><td>
                    <%
                        Application.Lock
                        Response.Write "<span class='success'>Lock acquired</span>"
                    %>
                </td></tr>
                <tr><td>Setting counter variable</td><td>
                    <%
                        Application("Counter") = Application("Counter") + 1
                        Response.Write "<span class='success'>Counter = " & Application("Counter") & "</span>"
                    %>
                </td></tr>
                <tr><td>Calling Application.Unlock()</td><td>
                    <%
                        Application.Unlock
                        Response.Write "<span class='success'>Lock released</span>"
                    %>
                </td></tr>
            </table>
        </div>

        <div class="section">
            <h3>Application.Contents Collection</h3>
            <p>Enumerating Application.Contents:</p>
            <table>
                <%
                    Dim item, contents
                    ' In ASP, you can iterate through Application.Contents
                    For Each item In Application.Contents
                        Response.Write "<tr><td>" & item & "</td><td><span class='value'>" & Application(item) & "</span></td></tr>"
                    Next
                %>
            </table>
        </div>

        <div class="section">
            <h3>Application.StaticObjects Collection</h3>
            <p>Note: StaticObjects should contain objects declared with &lt;OBJECT&gt; tags in Global.asa</p>
            <table>
                <%
                    ' Check if static objects exist
                    Dim staticObjCount
                    staticObjCount = 0
                    For Each item In Application.StaticObjects
                        staticObjCount = staticObjCount + 1
                        Response.Write "<tr><td>" & item & "</td><td><span class='info'>Object</span></td></tr>"
                    Next
                    
                    If staticObjCount = 0 Then
                        Response.Write "<tr><td colspan='2'><span class='error'>No static objects defined in Global.asa</span></td></tr>"
                    End If
                %>
            </table>
        </div>

        <h2>Session Object Tests</h2>

        <div class="section">
            <h3>Session Variables from Global.asa</h3>
            <table>
                <tr><td>Global_SessionStart_Time</td><td>
                    <%
                        If Not IsEmpty(Session("Global_SessionStart_Time")) And Not IsNull(Session("Global_SessionStart_Time")) Then
                            Response.Write "<span class='success'>Set: " & Session("Global_SessionStart_Time") & "</span>"
                        Else
                            Response.Write "<span class='error'>NOT SET (Session_OnStart may not have run)</span>"
                        End If
                    %>
                </td></tr>
                <tr><td>Global_SessionStart_Msg</td><td>
                    <%
                        If Not IsEmpty(Session("Global_SessionStart_Msg")) And Not IsNull(Session("Global_SessionStart_Msg")) Then
                            Response.Write "<span class='success'>" & Session("Global_SessionStart_Msg") & "</span>"
                        Else
                            Response.Write "<span class='error'>NOT SET</span>"
                        End If
                    %>
                </td></tr>
            </table>
        </div>

        <div class="section">
            <h3>Session.SessionID</h3>
            <table>
                <tr><td>SessionID Value</td><td><span class='value'><%=Session.SessionID%></span></td></tr>
                <tr><td>SessionID is string</td><td><%If IsObject(Session.SessionID) Then Response.Write "<span class='error'>NO (is object)</span>" Else Response.Write "<span class='success'>YES</span>" End If%></td></tr>
            </table>
        </div>

        <div class="section">
            <h3>Session.TimeOut Property</h3>
            <table>
                <tr><td>Current Session.TimeOut</td><td><span class='value'><%=Session.TimeOut%> minutes</span></td></tr>
                <tr><td>Set new timeout (30 min)</td><td>
                    <%
                        Session.TimeOut = 30
                        Response.Write "<span class='success'>New TimeOut: " & Session.TimeOut & " minutes</span>"
                    %>
                </td></tr>
            </table>
        </div>

        <div class="section">
            <h3>Session.Contents Collection</h3>
            <p>Enumerating Session.Contents:</p>
            <table>
                <%
                    Dim sessItem
                    Dim sessCount
                    sessCount = 0
                    For Each sessItem In Session.Contents
                        sessCount = sessCount + 1
                        Response.Write "<tr><td>" & sessItem & "</td><td><span class='value'>" & Session(sessItem) & "</span></td></tr>"
                    Next
                    
                    If sessCount = 0 Then
                        Response.Write "<tr><td colspan='2'><span class='error'>No session variables set</span></td></tr>"
                    End If
                %>
            </table>
        </div>

        <h2>Global.asa File Check</h2>

        <div class="section">
            <%
                ' Try to check if global.asa exists
                Dim fso
                Set fso = Server.CreateObject("Scripting.FileSystemObject")
                Dim globalAsaPath
                globalAsaPath = Server.MapPath("/global.asa")
                
                If fso.FileExists(globalAsaPath) Then
                    Response.Write "<span class='success'>✓ Global.asa file exists</span><br/>"
                    
                    ' Try to read and show first 500 chars
                    Dim file, content
                    Set file = fso.OpenTextFile(globalAsaPath, 1)
                    content = file.Read(500)
                    file.Close
                    
                    Response.Write "<p>Content preview:</p>"
                    Response.Write "<pre>" & Replace(Replace(content, "<", "&lt;"), ">", "&gt;") & "...</pre>"
                Else
                    Response.Write "<span class='error'>✗ Global.asa file NOT found</span><br/>"
                    Response.Write "Expected path: " & globalAsaPath
                End If
                
                Set fso = Nothing
            %>
        </div>

        <h2>Summary</h2>

        <div class="section">
            <%
                Dim testsPassed, testsFailed
                testsPassed = 0
                testsFailed = 0
                
                ' Check Application variables
                If Not IsEmpty(Application("Global_AppStart_Time")) Then
                    testsPassed = testsPassed + 1
                Else
                    testsFailed = testsFailed + 1
                End If
                
                ' Check Session variables
                If Not IsEmpty(Session("Global_SessionStart_Time")) Then
                    testsPassed = testsPassed + 1
                Else
                    testsFailed = testsFailed + 1
                End If
                
                ' Check Application.Lock/Unlock
                If IsNumeric(Application("Counter")) Then
                    testsPassed = testsPassed + 1
                Else
                    testsFailed = testsFailed + 1
                End If
                
                ' Check Session.TimeOut
                If IsNumeric(Session.TimeOut) Then
                    testsPassed = testsPassed + 1
                Else
                    testsFailed = testsFailed + 1
                End If
                
                Response.Write "<strong>Tests Passed: <span class='success'>" & testsPassed & "</span></strong><br/>"
                Response.Write "<strong>Tests Failed: <span class='error'>" & testsFailed & "</span></strong><br/>"
                Response.Write "<br/>"
                
                If testsFailed = 0 Then
                    Response.Write "<span class='success'>✓ All Global.asa tests passed!</span>"
                Else
                    Response.Write "<span class='error'>✗ Some tests failed. Check Global.asa is properly configured.</span>"
                End If
            %>
        </div>

        <footer style="margin-top: 40px; padding-top: 20px; border-top: 2px solid #0066cc; text-align: center;">
            <p><strong>G3 AxonASP</strong> - Global.asa Support v1.0</p>
        </footer>
    </div>
</body>
</html>
