<%
' Test data preservation across requests
Dim testValue

' On first request, set session value
' On second request, verify it persists

Session("testkey") = "testvalue123"
Session("counter") = CLng(Session("counter")) + 1

Response.Write("<h1>Session Test</h1>" & vbCrLf)
Response.Write("<p>Counter: " & Session("counter") & "</p>" & vbCrLf)
Response.Write("<p>Test Value: " & Session("testkey") & "</p>" & vbCrLf)
Response.Write("<p>Session ID: " & Session("sessionid") & "</p>" & vbCrLf)

Response.Write("<hr />" & vbCrLf)
Response.Write("<p><a href='?refresh=1'>Refresh</a> to test session persistence</p>" & vbCrLf)
Response.Write("<p>If Counter increases and Test Value persists, sessions are working!</p>" & vbCrLf)
%>
