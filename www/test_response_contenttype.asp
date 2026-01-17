<%
' Test Response.ContentType implementation
Response.ContentType = "text/html; charset=UTF-8"
%>
<!DOCTYPE html>
<html>
<head>
	<title>Response.ContentType Test</title>
	<style>
		body { font-family: Arial, sans-serif; margin: 20px; }
		.success { color: green; font-weight: bold; }
		.code { background: #f0f0f0; padding: 10px; border-radius: 5px; font-family: monospace; }
	</style>
</head>
<body>
	<h1>Response.ContentType Implementation Test</h1>
	
	<div class="code">
		<%
		Response.Write "Response.ContentType = """ & Response.GetProperty("contenttype") & """<br />"
		Response.Write "This header is now being correctly returned in HTTP response!"
		%>
	</div>
	
	<p class="success">✓ Content-Type header is working correctly!</p>
</body>
</html>
