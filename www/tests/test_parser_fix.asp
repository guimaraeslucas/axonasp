<%
response.write "Testing parsing of response.end without calling it"
if false then
	response.end
end if
response.write "<br>Success - parser handled response.end correctly!"
%>
