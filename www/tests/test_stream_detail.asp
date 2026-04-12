<%
' More detailed debug
Option Explicit

Response.Write("<h1>Detailed Stream Debug</h1>")

Dim objStream, data

Response.Write("<h2>Test 1: Check if LoadFromFile works</h2>")
Set objStream = server.CreateObject("ADODB.Stream")
Response.Write("1. Created stream<br>")

' Try calling open first
objStream.Open
Response.Write("2. Called Open, state=" & objStream.state & "<br>")

' Set type
objStream.type = 1
Response.Write("3. Set type=1, state=" & objStream.state & "<br>")

' Load file
objStream.LoadFromFile(server.mappath("demo_file.txt"))
Response.Write("4. LoadFromFile done, state=" & objStream.state & "<br>")

' Try read
data = objStream.Read()
Response.Write("5. Read returned " & Len(data) & " bytes<br>")
Response.Write("6. Content: " & Server.HTMLEncode(CStr(data)) & "<br>")

objStream.Close()
Set objStream = Nothing

Response.Write("<h2>Test 2: Without calling Open</h2>")
Set objStream = server.CreateObject("ADODB.Stream")
objStream.type = 1
Response.Write("1. Set type=1, state=" & objStream.state & "<br>")
objStream.LoadFromFile(server.mappath("demo_file.txt"))
Response.Write("2. LoadFromFile done, state=" & objStream.state & "<br>")

data = objStream.Read()
Response.Write("3. Read returned " & Len(data) & " bytes<br>")

objStream.Close()
Set objStream = Nothing

Response.Write("<p>Done</p>")
%>
