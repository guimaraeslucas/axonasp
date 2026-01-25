<%
' Simple test of the fix
select case "test"
	case "php"
		response.write "PHP case: " : response.end
	case "test"
		response.write "Test case matched"
end select
%>
