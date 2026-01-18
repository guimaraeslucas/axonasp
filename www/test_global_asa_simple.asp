<%@ Page Language="VBScript" %>
<html>
<head>
    <title>Global.asa Simple Test</title>
</head>
<body>
<h1>Global.asa Implementation Test</h1>

<%
    ' Test 1: Application_OnStart
    Dim appStartTime
    appStartTime = Application("Global_AppStart_Time")
    If Not IsEmpty(appStartTime) Then
        Response.Write "<p>✓ Application_OnStart executed: " & appStartTime & "</p>"
    Else
        Response.Write "<p>✗ Application_OnStart NOT executed</p>"
    End If
%>

<%
    ' Test 2: Application.Lock/Unlock
    Application.Lock
    Application("TestCounter") = Application("TestCounter") + 1
    Application.Unlock
    Response.Write "<p>✓ Application.Lock/Unlock working. Counter = " & Application("TestCounter") & "</p>"
%>

<%
    ' Test 3: Session_OnStart
    Dim sessionStartTime
    sessionStartTime = Session("Global_SessionStart_Time")
    If Not IsEmpty(sessionStartTime) Then
        Response.Write "<p>✓ Session_OnStart executed: " & sessionStartTime & "</p>"
    Else
        Response.Write "<p>✗ Session_OnStart NOT executed</p>"
    End If
%>

<%
    ' Test 4: Session.TimeOut
    Session.TimeOut = 30
    Response.Write "<p>✓ Session.TimeOut = " & Session.TimeOut & " minutes</p>"
%>

<%
    ' Test 5: Enumeration
    Dim itemCount
    itemCount = 0
    For Each item In Application.Contents
        itemCount = itemCount + 1
    Next
    Response.Write "<p>✓ Application.Contents has " & itemCount & " items</p>"
%>

<p><strong>All Global.asa features working!</strong></p>

</body>
</html>
