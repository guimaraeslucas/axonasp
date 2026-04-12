<%
Response.LCID = 1046  ' Portuguese Brazil

Response.Write "=== Response.LCID Test ===" & vbCrLf & vbCrLf

Response.Write "Response.LCID = " & Response.LCID & vbCrLf
Response.Write "Now() = " & Now() & vbCrLf
Response.Write "FormatDateTime(Now(), vbLongTime) = " & FormatDateTime(Now(), vbLongTime) & vbCrLf
Response.Write "FormatDateTime(Now(), vbShortDate) = " & FormatDateTime(Now(), vbShortDate) & vbCrLf
Response.Write "FormatDateTime(Now(), vbGeneralDate) = " & FormatDateTime(Now(), vbGeneralDate) & vbCrLf
Response.Write vbCrLf & "Expected for Portuguese (pt-BR):" & vbCrLf
Response.Write "- vbLongTime format: HH:MM:SS (24-hour, no AM/PM)" & vbCrLf
Response.Write "- vbShortDate format: DD/MM/YYYY" & vbCrLf
%>
