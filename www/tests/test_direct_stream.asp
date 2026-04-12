<%
' Direct usage of ADODB.Stream as in user's code
Option Explicit

Response.Write("<h1>Direct ADODB.Stream Test</h1>")

' Test 1: Text mode (type 2)
Response.Write("<h2>Test 1: Read Text File</h2>")
Dim objStream1
Set objStream1 = server.CreateObject("ADODB.Stream")

objStream1.CharSet = "utf-8"
objStream1.Open
objStream1.type = 2 'adTypeText
objStream1.LoadFromFile(server.mappath("demo_file.txt"))

Dim textContent
textContent = objStream1.ReadText()
Dim textSize
textSize = objStream1.size

Response.Write("Text content length: " & Len(textContent) & "<br>")
Response.Write("Stream size: " & textSize & "<br>")
Response.Write("Content: " & Server.HTMLEncode(textContent) & "<br>")

objStream1.Close()
Set objStream1 = Nothing

' Test 2: Binary mode (type 1)
Response.Write("<h2>Test 2: Read Binary File</h2>")
Dim objStream2
Set objStream2 = server.CreateObject("ADODB.Stream")

objStream2.Open
objStream2.type = 1 'adTypeBinary
objStream2.LoadFromFile(server.mappath("demo_file.txt"))

Dim binContent
binContent = objStream2.Read()
Dim binSize
binSize = objStream2.size

Response.Write("Binary content length: " & Len(binContent) & "<br>")
Response.Write("Stream size: " & binSize & "<br>")

objStream2.Close()
Set objStream2 = Nothing

Response.Write("<p><strong>✓ Both text and binary reading working!</strong></p>")
%>
