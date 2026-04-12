<%
' Test: single-line vs multi-line if
dim x
x = 2

' Multi-line IF
if x = 2 then
  Response.Write "Multi-line: x is 2"
else
  Response.Write "Multi-line: x is not 2"
end if

Response.Write "<br />"

' Single-line IF  
if x = 2 then Response.Write "Single-line: x is 2" else Response.Write "Single-line: x is not 2"

Response.Write "<br />"
%>
