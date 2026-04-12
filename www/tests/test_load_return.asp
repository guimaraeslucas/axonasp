<%
' Test if LoadFromFile returns error or not
Option Explicit

Dim objStream
Set objStream = server.CreateObject("ADODB.Stream")

' Try to load file
Dim result
result = objStream.LoadFromFile(server.mappath("demo_file.txt"))
Response.Write("LoadFromFile returned: " & result & "<br>")
Response.Write("Then State: " & objStream.state & "<br>")

objStream.Close()
Set objStream = Nothing
%>
