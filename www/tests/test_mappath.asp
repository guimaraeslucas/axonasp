<%
' Test MapPath
Option Explicit

Dim path
path = server.mappath("demo_file.txt")
Response.Write("MapPath result: " & path & "<br>")

' Try to manually check if file exists
Response.Write("Now creating stream...<br>")

Dim objStream
Set objStream = server.CreateObject("ADODB.Stream")

Response.Write("Before LoadFromFile - Size: " & objStream.size & "<br>")
Response.Write("Calling LoadFromFile with path: " & server.mappath("demo_file.txt") & "<br>")

objStream.LoadFromFile(server.mappath("demo_file.txt"))

Response.Write("After LoadFromFile - Size: " & objStream.size & "<br>")
Response.Write("Position: " & objStream.position & "<br>")

objStream.Close()
Set objStream = Nothing
%>
