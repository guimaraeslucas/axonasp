<%
Response.Write "<h1>FormatDateTime Test</h1>"

d = Now()
Response.Write "Date used: " & d & "<br><br>"

Response.Write "vbGeneralDate (0): " & FormatDateTime(d, vbGeneralDate) & "<br>"
Response.Write "vbLongDate (1): " & FormatDateTime(d, vbLongDate) & "<br>"
Response.Write "vbShortDate (2): " & FormatDateTime(d, vbShortDate) & "<br>"
Response.Write "vbLongTime (3): " & FormatDateTime(d, vbLongTime) & "<br>"
Response.Write "vbShortTime (4): " & FormatDateTime(d, vbShortTime) & "<br>"

Response.Write "<br>Numeric values check:<br>"
Response.Write "0: " & FormatDateTime(d, 0) & "<br>"
Response.Write "1: " & FormatDateTime(d, 1) & "<br>"
Response.Write "2: " & FormatDateTime(d, 2) & "<br>"
Response.Write "3: " & FormatDateTime(d, 3) & "<br>"
Response.Write "4: " & FormatDateTime(d, 4) & "<br>"
%>
