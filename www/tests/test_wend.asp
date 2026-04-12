<%
Option Explicit
Response.Write("Testing WEND Keyword" & vbCrLf & vbCrLf)

Dim counter
counter = 0

While counter < 5
    Response.Write("Counter: " & counter & vbCrLf)
    counter = counter + 1
Wend

Response.Write(vbCrLf & "WEND test passed!" & vbCrLf)
%>
