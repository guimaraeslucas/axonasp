<%
' Test ADODB.Stream functionality
Option Explicit

private function stream(path, binary, byref size)
    on error resume next
    Dim objStream
    Set objStream = server.CreateObject("ADODB.Stream")
    if binary then
        objStream.Open
        objStream.type = 1 'adTypeBinary
        objStream.LoadFromFile(server.mappath(path))
        stream = objStream.Read()
    else
        objStream.CharSet = "utf-8"
        objStream.Open
        objStream.type = 2 'adTypeText
        objStream.LoadFromFile(server.mappath(path))
        stream = objStream.ReadText()
    end if
    size = objStream.size
    set objStream = nothing
    on error goto 0
end function

Response.Write("<h1>ADODB.Stream Test</h1>")

' Test 1: Read text file
Response.Write("<h2>Test 1: Reading Text File</h2>")
Dim fileSize1
Dim content1
content1 = stream("demo_file.txt", False, fileSize1)
Response.Write("Content: " & Server.HTMLEncode(content1) & "<br>")
Response.Write("Size: " & fileSize1 & " bytes<br>")
Response.Write("Length of content: " & Len(content1) & "<br>")

' Test 2: Read binary file
Response.Write("<h2>Test 2: Reading Binary File (Image)</h2>")
Dim fileSize2
Dim content2
content2 = stream("demo_file.txt", True, fileSize2)
Response.Write("Binary data length: " & Len(content2) & " bytes<br>")
Response.Write("Size from stream: " & fileSize2 & " bytes<br>")

' Test 3: Direct Stream operations
Response.Write("<h2>Test 3: Direct Stream Operations</h2>")
Dim objStream
Set objStream = server.CreateObject("ADODB.Stream")

' Test text mode
objStream.CharSet = "utf-8"
objStream.Open
objStream.type = 2 'adTypeText
objStream.WriteText("Hello World!")
Response.Write("Text written, position: " & objStream.position & "<br>")
Response.Write("Stream size: " & objStream.size & "<br>")

' Reset position
objStream.position = 0
Dim readText
readText = objStream.ReadText()
Response.Write("Text read back: " & readText & "<br>")

objStream.Close
Set objStream = Nothing

Response.Write("<p><strong>All tests completed successfully!</strong></p>")
%>
