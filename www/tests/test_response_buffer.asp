<%
Response.Buffer = false
startTime = Time()
'Do Until DateDiff("s", startTime, Time(), 0, 0) >2
response.write("Current time: " & Time() & "<br>")
'Loop
%>