<%
Option Explicit

Response.Write("<h1>Simple Stream Test</h1>")

Dim objStream
Set objStream = server.CreateObject("ADODB.Stream")

' Don't call Open, just load
objStream.LoadFromFile(server.mappath("demo_file.txt"))

Dim data
data = objStream.Read()

Response.Write("Bytes read: " & Len(data) & "<br>")
Response.Write("Content: " & Server.HTMLEncode(CStr(data)) & "<br>")

objStream.Close()
Set objStream = Nothing
%>
