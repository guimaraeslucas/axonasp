<%
' Simple ADODB.Stream test
Option Explicit

Response.Write("<h1>ADODB.Stream Direct Test</h1>")

' Create and test stream directly
Dim objStream
Set objStream = server.CreateObject("ADODB.Stream")

' Test binary read from file
Response.Write("<h2>Test: Binary Read from File</h2>")
objStream.type = 1 'adTypeBinary
objStream.LoadFromFile(server.mappath("demo_file.txt"))

Response.Write("Stream state: " & objStream.state & "<br>")
Response.Write("Stream type: " & objStream.type & "<br>")
Response.Write("Stream size: " & objStream.size & "<br>")

Dim binData
binData = objStream.Read()
Response.Write("Data length: " & Len(binData) & "<br>")

objStream.Close()

' Test text read from file
Response.Write("<h2>Test: Text Read from File</h2>")
Dim objStream2
Set objStream2 = server.CreateObject("ADODB.Stream")

objStream2.CharSet = "utf-8"
objStream2.type = 2 'adTypeText
objStream2.LoadFromFile(server.mappath("demo_file.txt"))

Response.Write("Stream state: " & objStream2.state & "<br>")
Response.Write("Stream type: " & objStream2.type & "<br>")
Response.Write("Stream size: " & objStream2.size & "<br>")

Dim textData
textData = objStream2.ReadText()
Response.Write("Text data: " & Server.HTMLEncode(textData) & "<br>")
Response.Write("Data length: " & Len(textData) & "<br>")

objStream2.Close()
Set objStream2 = Nothing

Set objStream = Nothing

Response.Write("<p><strong>Tests completed!</strong></p>")
%>
