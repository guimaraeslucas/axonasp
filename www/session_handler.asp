<%
Dim action
action = Request.QueryString("action")

If action = "set" Then
    ' Set initial session values
    Session("username") = "TestUser"
    Session("counter") = 1
    Response.Write "Session Set:" & vbCrLf
    Response.Write "  Username: " & Session("username") & vbCrLf
    Response.Write "  Counter: " & Session("counter") & vbCrLf
    Response.Write "  SessionID: " & Session.SessionID
    
ElseIf action = "increment" Then
    ' Increment counter
    Dim count
    count = Session("counter")
    If count = "" Or IsNull(count) Then
        count = 0
    End If
    count = count + 1
    Session("counter") = count
    
    Response.Write "Counter Incremented:" & vbCrLf
    Response.Write "  Old Value: " & (count - 1) & vbCrLf
    Response.Write "  New Value: " & count & vbCrLf
    Response.Write "  SessionID: " & Session.SessionID
    
ElseIf action = "get" Then
    ' Retrieve session data
    Response.Write "Session Data Retrieved:" & vbCrLf
    Response.Write "  Username: " & Session("username") & vbCrLf
    Response.Write "  Counter: " & Session("counter") & vbCrLf
    Response.Write "  SessionID: " & Session.SessionID
    
Else
    Response.Write "Unknown action: " & action
End If
%>
