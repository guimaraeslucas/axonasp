<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Buffer				= true
Response.ContentType		= "text/html"
Response.CacheControl		= "no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires			= -1
Response.ExpiresAbsolute	= Now()-1
%>
<!doctype html>
<html lang="en">
  <head>   
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-T3c6CoIi6uLrA9TneNEoa7RxnatzjcDSCmG1MXxSR1GAsXEV/Dwwykc2MPK8M2HN" crossorigin="anonymous">	
	<title>Media settings</title>
	<style>
		body {
			scroll-behavior:smooth;
			padding:20px;					
		}
	</style>
  </head>
  <body>  
	
	<div class="asplForm" id="upload"></div>
	<div class="asplForm" id="media"></div>		
		
	<!-- jQuery  & aspLite JS -->
	<script src="asplite/jquery.js"></script>
	
	<script>var aspLiteAjaxHandler='bs_itemmedia_handler.asp?iId=<%=server.htmlencode(request.querystring("iId"))%>'</script>
	
	<script src="asplite/asplite.js"></script>
	
  </body>
</html>

