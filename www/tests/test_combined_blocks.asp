<%
' Test: Code blocks separated by HTML
dim counter
counter = 0
%>
<p>Start</p>
<%
counter = counter + 1
%>
<p>Counter is: <%=counter%></p>
<%
counter = counter + 1
%>
<p>Counter is now: <%=counter%></p>
<%
if counter = 2 then
%>
<p>Counter equals 2</p>
<%
else
%>
<p>Counter does not equal 2</p>
<%
end if
%>
<p>End</p>
