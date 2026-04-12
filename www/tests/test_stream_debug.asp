<%
' Debug ADODB.Stream test
Option Explicit

Response.Write("<h1>ADODB.Stream Debug Test</h1>")

Dim objStream
Set objStream = server.CreateObject("ADODB.Stream")
Response.Write("Created stream<br>")

objStream.type = 1
Response.Write("Set type to 1 (binary)<br>")

objStream.LoadFromFile(server.mappath("demo_file.txt"))
Response.Write("Called LoadFromFile<br>")

' Check state immediately
Response.Write("Checking properties:<br>")
Response.Write("- State: " & objStream.state & "<br>")
Response.Write("- Type: " & objStream.type & "<br>")
Response.Write("- Size: " & objStream.size & "<br>")

' Try to read
Dim data
data = objStream.Read()
Response.Write("- Data read, length: " & Len(data) & "<br>")

objStream.Close()
Set objStream = Nothing

Response.Write("<p>Done</p>")
%>
