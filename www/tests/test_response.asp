<%
Response.Write "<h3>Response Object Test</h3>"

Response.ContentType "text/html"
Response.CacheControl "No-Cache"
Response.Charset "utf-8"
Response.Status "200 OK"
Response.Buffer True

Response.AddHeader "X-Axon-Test", "response"
Response.AppendToLog "Response object test log"

Response.Cookies "sid", "abc123"
Response.Cookies "sid", "Domain", "localhost"
Response.Cookies "sid", "Path", "/"
Response.Cookies "sid", "Secure", "False"
Response.Cookies "sid", "HttpOnly", "True"

Response.Write "ContentType: " & Response.ContentType() & "<br>"
Response.Write "CacheControl: " & Response.CacheControl() & "<br>"
Response.Write "Status: " & Response.Status() & "<br>"
Response.Write "Cookie sid: " & Response.Cookies("sid") & "<br>"
Response.Write "Cookie sid Domain: " & Response.Cookies("sid", "Domain") & "<br>"
Response.Write "Cookie sid Domain (chain): " & Response.Cookies("sid").Domain & "<br>"
Response.Write "ClientConnected: " & Response.IsClientConnected() & "<br>"

Response.Flush
%>
