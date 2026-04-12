<%
' Test to check if timezone is being used
Dim currentHour, expectedHour

currentHour = Hour(Now())
expectedHour = 17  ' São Paulo should be UTC - 3, so 20 - 3 = 17

Response.Write "Current Hour: " & currentHour & "<br>"
Response.Write "Expected Hour (SP UTC-3): " & expectedHour & "<br>"

If currentHour = expectedHour Then
    Response.Write "✓ TIMEZONE WORKING! Hour is correct.<br>"
ElseIf currentHour = 20 Then
    Response.Write "✗ TIMEZONE NOT WORKING! Using UTC hour (20).<br>"
Else
    Response.Write "? UNEXPECTED HOUR! May be using different timezone.<br>"
End If

Response.Write "<br>"
Response.Write "Now() = " & Now() & "<br>"
Response.Write "FormatDateTime(Now, 3) = " & FormatDateTime(Now(), 3) & "<br>"
%>
