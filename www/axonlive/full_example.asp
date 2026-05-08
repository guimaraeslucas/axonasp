<%@ Language="VBScript" %>
<%
'
' G3AxonLive Full Capability Showcase
' Copyright (C) 2026 G3pix Ltda. All rights reserved.
'
' This example demonstrates all major features of the G3AxonLive framework:
' 1. Full HTML Replacement (RegisterComponent)
' 2. Granular Property Manipulation (Proxy Objects)
' 3. Style and CSS Class management
' 4. DOM Attribute manipulation
' 5. Server-triggered client actions (Timer, Redirect, Trigger)
' 6. Event Arguments
'

Dim AxonLive
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

Dim sessionID : sessionID = Session.SessionID

' --- Load Persisted State ---
Dim clickCount
clickCount = AxonLive.GetComponentProperty(sessionID, "main", "clicks")
If IsEmpty(clickCount) Or clickCount = "" Then clickCount = 0 Else clickCount = CLng(clickCount)

Dim lastAction
lastAction = AxonLive.GetComponentProperty(sessionID, "main", "lastAction")
If IsEmpty(lastAction) Then lastAction = "None"

' --- Handle Async Events ---
If AxonLive.IsAsyncRequest Then
    Dim compID : compID = AxonLive.EventComponentID
    Dim evtName : evtName = AxonLive.EventName
    
    ' Get proxies for common elements
    Dim statusPill, btnAction, txtLog, btnGhost
    Set statusPill = AxonLive.GetComponent("statusPill")
    Set btnAction = AxonLive.GetComponent("btnAction")
    Set txtLog = AxonLive.GetComponent("txtLog")
    Set btnGhost = AxonLive.GetComponent("btnGhost")

    Select Case compID
        Case "btnAction"
            clickCount = clickCount + 1
            lastAction = "Button Clicked"
            
            ' Granular Update: Set property
            statusPill.value = "Active (Clicks: " & clickCount & ")"
            
            ' Granular Update: Set Style & Class
            If clickCount Mod 2 = 0 Then
                statusPill.SetStyle "background-color", "#3366cc"
                statusPill.AddClass "status-v"
                statusPill.RemoveClass "status-x"
            Else
                statusPill.SetStyle "background-color", "#cc3300"
                statusPill.AddClass "status-x"
                statusPill.RemoveClass "status-v"
            End If
            
            ' Granular Update: Set Attribute & Title
            btnAction.SetAttribute "data-count", CStr(clickCount)
            btnAction.AddTitle "You have clicked " & clickCount & " times"
            
            ' Append to log using existing value (Persistence Test)
            Dim currentLog
            currentLog = AxonLive.GetComponentProperty(sessionID, "txtLog", "val")
            txtLog.value = currentLog & vbCrLf & "[" & Now & "] Action performed. Total: " & clickCount

        Case "btnGhost"
            ' Demonstration of "Trigger" action
            ' This button will "trigger" btnAction after 1 second
            lastAction = "Ghost Trigger Scheduled"
            AxonLive.SetTimer "btnAction", "onclick", 1000
            btnGhost.disabled = True
            btnGhost.value = "Triggering in 1s..."
            
        Case "btnRedirect"
            ' Demonstration of "Redirect" action
            AxonLive.Redirect "https://g3pix.com.br/axonasp"
            
        Case "btnTrigger"
            ' Demonstration of immediate "Trigger" (Server-to-Client-to-Server)
            ' This forces the browser to immediately fire btnAction.onclick
            AxonLive.Trigger "btnAction", "onclick"
            lastAction = "Remote Trigger Fired"

        Case "btnArgs"
            ' Demonstration of Event Arguments
            Dim stepVal
            stepVal = AxonLive.GetEventArg("step")
            If stepVal = "" Then stepVal = 1 Else stepVal = CLng(stepVal)
            
            clickCount = clickCount + stepVal
            lastAction = "Custom Step Addition: " & stepVal
            
            ' Update everything
            statusPill.value = "Updated by " & stepVal & " (Total: " & clickCount & ")"
            txtLog.value = "Arguments received. Step=" & stepVal

        Case "timer1"
            ' Demonstration of an auto-repeating server timer
            clickCount = clickCount + 1
            lastAction = "Auto Timer Tick"
            
            statusPill.value = "Timer Tick: " & clickCount
            ' Re-schedule the timer
            AxonLive.SetTimer "timer1", "ontimer", 5000 
    End Select

    ' Persist state
    Call AxonLive.SetComponentProperty(sessionID, "main", "clicks", CStr(clickCount))
    Call AxonLive.SetComponentProperty(sessionID, "main", "lastAction", lastAction)
    Call AxonLive.SetComponentProperty(sessionID, "txtLog", "val", txtLog.value)

    ' Full Replacement Demo: Update the "Last Action" card
    AxonLive.RegisterComponent "cardAction", _
        "<div id=""cardAction"" class=""card"">" & _
        "<h3>Last Server Action</h3>" & _
        "<p class=""pill pill-primary"">" & Server.HTMLEncode(lastAction) & "</p>" & _
        "<p>Timestamp: " & Now & "</p>" & _
        "</div>"

    AxonLive.EndAsyncResponse()
End If
%>
<!DOCTYPE html>
<html>
<head>
    <title>AxonLive Full Capability Demo</title>
    <link rel="stylesheet" href="/css/axonasp.css">
    <style>
        .demo-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        #txtLog { width: 100%; height: 150px; font-family: monospace; font-size: 11px; margin-top: 10px; }
        .status-pill-big { font-size: 18px; padding: 15px; text-align: center; font-weight: bold; border-radius: 8px; color: #fff; background: #808080; transition: all 0.3s; }
    </style>
</head>
<body>

<div id="header">
    <span style="color:#fff; font-family:Tahoma,Verdana,Arial,sans-serif; font-size:18px; font-weight:bold; line-height:60px; padding-left:18px;">
        G3AxonLive &mdash; Full Capability Demonstration
    </span>
</div>

<div id="main-container">
    <div id="content">
        
        <div class="info-banner">
            This page demonstrates the full spectrum of AxonLive's reactive capabilities, from granular property updates to server-triggered client orchestration.
        </div>

        <div class="demo-grid">
            <!-- Left Column: Controls -->
            <div class="col">
                <div class="card">
                    <h3>Interactive Controls</h3>
                    <p>These buttons use different methods to interact with the server.</p>
                    
                    <div class="actions-row">
                        <button id="btnAction" class="btn btn-primary" 
                                data-g3al-id="btnAction" data-g3al-event="click" data-g3al-event-name="onclick">
                            Primary Action
                        </button>
                        
                        <button id="btnArgs" class="btn btn-secondary" 
                                data-g3al-id="btnArgs" data-g3al-event="click" data-g3al-event-name="btnArgs"
                                data-g3al-arg-step="5">
                            Add +5 (Args)
                        </button>
                    </div>

                    <div class="actions-row" style="margin-top:10px;">
                        <button id="btnGhost" class="btn btn-info" 
                                data-g3al-id="btnGhost" data-g3al-event="click" data-g3al-event-name="btnGhost">
                            Ghost (Delayed Trigger)
                        </button>
                        
                        <button id="btnTrigger" class="btn btn-warning" 
                                data-g3al-id="btnTrigger" data-g3al-event="click" data-g3al-event-name="btnTrigger">
                            Remote Trigger
                        </button>
                    </div>

                    <div class="actions-row" style="margin-top:10px;">
                        <button id="btnRedirect" class="btn btn-download" 
                                data-g3al-id="btnRedirect" data-g3al-event="click" data-g3al-event-name="btnRedirect">
                            External Redirect
                        </button>
                    </div>
                </div>

                <div class="card">
                    <h3>Server Log</h3>
                    <textarea id="txtLog" readonly>Page Loaded: <%=Now%></textarea>
                </div>
            </div>

            <!-- Right Column: Display -->
            <div class="col">
                <div id="statusPill" class="status-pill-big">
                    Ready to start...
                </div>

                <div id="cardAction" class="card">
                    <h3>Last Server Action</h3>
                    <p class="pill pill-primary"><%=lastAction%></p>
                    <p>Timestamp: <%=Now%></p>
                </div>

                <div class="card">
                    <h3>Framework Capabilities</h3>
                    <ul class="treeview">
                        <li class="folder">Granular Updates (O(1))
                            <ul class="submenu">
                                <li class="file">Property Set (value, disabled, checked)</li>
                                <li class="file">Style Manipulation (SetStyle)</li>
                                <li class="file">Class Manipulation (Add/RemoveClass)</li>
                                <li class="file">Attribute Manipulation</li>
                            </ul>
                        </li>
                        <li class="folder">Orchestration
                            <ul class="submenu">
                                <li class="file">Server-to-Client Timers</li>
                                <li class="file">Server-to-Client Redirects</li>
                                <li class="file">Server-to-Client Event Triggers</li>
                            </ul>
                        </li>
                        <li class="file">Persistence (Cross-request state)</li>
                    </ul>
                </div>
            </div>
        </div>

        <!-- Hidden Timer Component -->
        <div id="timer1" data-g3al-id="timer1" data-g3al-event="timer" data-g3al-event-name="timer1" style="display:none"></div>

    </div>
</div>

<div id="status-bar">
    Session: <%=Session.SessionID%> | Total Clicks: <%=clickCount%>
</div>

<script src="/axonlive/g3axonlive.js"></script>
<script>
    G3AxonLive.init('<%=Server.HTMLEncode(Session.SessionID)%>');
    
    // Start a background timer tick from the client side initially
    setTimeout(function() {
        G3AxonLive.trigger('timer1', 'ontimer');
    }, 5000);
</script>

</body>
</html>
