<%
Response.Write "<h3>Server Object Test</h3>"

Response.Write "HTMLEncode: " & Server.HTMLEncode("<b>Axon</b>") & "<br>"
Response.Write "URLEncode: " & Server.URLEncode("a b&c") & "<br>"
Response.Write "URLPathEncode: " & Server.URLPathEncode("folder name/file.asp") & "<br>"
Response.Write "MapPath(/index.asp): " & Server.MapPath("/index.asp") & "<br>"

Response.Write "ScriptTimeout (before): " & Server.ScriptTimeout() & "<br>"
Server.ScriptTimeout 120
Response.Write "ScriptTimeout (after): " & Server.ScriptTimeout() & "<br>"

Server.CreateObject "ADODB.Connection"
Response.Write "LastError Number: " & Server.GetLastError("Number") & "<br>"
Response.Write "LastError Description: " & Server.GetLastError("Description") & "<br>"
%>
