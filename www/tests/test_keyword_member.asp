<%
' Test that keywords can be used as member names after a dot
Response.Write("Before calling response.end<br>")
if 1=2 then
	Response.end
end if
Response.Write("After the conditional (should appear)")
%>