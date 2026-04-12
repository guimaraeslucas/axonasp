<%
Response.Write "<h3>Request Object Test</h3>"

Response.Write "QueryString(name): " & Request.QueryString("name") & "<br>"
Response.Write "Form(data): " & Request.Form("data") & "<br>"
Response.Write "Cookie(profile): " & Request.Cookies("profile") & "<br>"
Response.Write "Cookie(profile.name): " & Request.Cookies("profile", "name") & "<br>"
Response.Write "ServerVariables(URL): " & Request.ServerVariables("URL") & "<br>"
Response.Write "TotalBytes: " & Request.TotalBytes() & "<br>"
Response.Write "Default Request(name): " & Request("name") & "<br>"

chunk = Request.BinaryRead(5)
Response.Write "BinaryRead(5): " & chunk & "<br>"
%>
