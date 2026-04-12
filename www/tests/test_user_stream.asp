<%
' Direct test of user's stream function
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

Response.Write("<h1>Testing User's Stream Function</h1>")

' Test 1: Read text file
Response.Write("<h2>Test 1: Read Text File</h2>")
Dim fileSize
Dim content
content = stream("demo_file.txt", False, fileSize)
Response.Write("Returned content length: " & Len(content) & "<br>")
Response.Write("File size from byref: " & fileSize & "<br>")
Response.Write("Content: " & Server.HTMLEncode(content) & "<br>")

' Test 2: Read binary file
Response.Write("<h2>Test 2: Read Binary File</h2>")
Dim binSize
Dim binContent
binContent = stream("demo_file.txt", True, binSize)
Response.Write("Binary content length: " & Len(binContent) & "<br>")
Response.Write("File size from byref: " & binSize & "<br>")

Response.Write("<p><strong>Done!</strong></p>")
%>
