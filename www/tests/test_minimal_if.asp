<%
' Minimal test with HTML blocks
dim x
x = 2
%>
<p>Before IF</p>
<%
if x = 2 then
%>
<p>In IF</p>
<%
else
%>
<p>In ELSE</p>
<%
end if
%>
<p>After IF</p>
