<%
Response.Write "<h3>Request Full Compatibility Test</h3>"

Response.Write "Default Request(name): " & Request("name") & "<br>"
Response.Write "QueryString(name): " & Request.QueryString("name") & "<br>"
Response.Write "QueryString.Count: " & Request.QueryString.Count & "<br>"
Response.Write "QueryString.Key(1): " & Request.QueryString.Key(1) & "<br>"
Response.Write "Form(data): " & Request.Form("data") & "<br>"
Response.Write "Cookies(profile): " & Request.Cookies("profile") & "<br>"
Response.Write "Cookies(profile,name): " & Request.Cookies("profile", "name") & "<br>"
Response.Write "ServerVariables(URL): " & Request.ServerVariables("URL") & "<br>"
Response.Write "TotalBytes: " & Request.TotalBytes & "<br>"
Response.Write "BinaryRead(4): " & Request.BinaryRead(4) & "<br>"
Response.Write "BinaryRead(100): " & Request.BinaryRead(100) & "<br>"
%>
