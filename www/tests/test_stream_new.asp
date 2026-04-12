<%
Option Explicit

Response.Write("<h1>Stream Debug</h1>")

Dim objStream
Set objStream = server.CreateObject("ADODB.Stream")
Response.Write("Created<br>")
Response.Write("State: " & objStream.state & "<br>")

objStream.LoadFromFile(server.mappath("demo_file.txt"))
Response.Write("After LoadFromFile<br>")
Response.Write("State: " & objStream.state & "<br>")
Response.Write("Size: " & objStream.size & "<br>")

Dim data
data = objStream.Read()
Response.Write("Read returned: " & Len(data) & " chars<br>")

objStream.Close()
Set objStream = Nothing
%>
