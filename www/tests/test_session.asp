<%
Response.Write "<h3>Session Object Test</h3>"

Session.Set "Counter", 1
Session.Set "UserName", "AxonASP"

Response.Write "SessionID: " & Session.SessionID() & "<br>"
Response.Write "Counter: " & Session.Get("counter") & "<br>"
Response.Write "UserName: " & Session.Get("USERNAME") & "<br>"
Response.Write "Exists(UserName): " & Session.Exists("username") & "<br>"
Response.Write "Count: " & Session.Count() & "<br>"

Session.Timeout 45
Response.Write "Timeout: " & Session.Timeout() & "<br>"

Session.Remove "username"
Response.Write "Exists(UserName) after Remove: " & Session.Exists("UserName") & "<br>"

Session.Removeall
Response.Write "Count after RemoveAll: " & Session.Count() & "<br>"
%>
