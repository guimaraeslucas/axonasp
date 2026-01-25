<%
Document.Write "Test 1: " & "<script>alert('XSS')</script>" & vbCrLf
Response.Write "Test 2: OK" & vbCrLf
%>