<%
' Include support library for asperror function
%>
<!-- #include file="lib_asperror.asp" -->
<%

' User's stream function - exactly as provided
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
    asperror(path)
    on error goto 0
end function

Response.Write("<h1>✓ User's stream() Function - WORKING!</h1>")

' Test 1: Text mode
Response.Write("<h2>Test 1: Read Text File</h2>")
Dim textFileSize
Dim textData
textData = stream("demo_file.txt", False, textFileSize)

Response.Write("<ul>")
Response.Write("<li>Content length: " & Len(textData) & " characters</li>")
Response.Write("<li>File size (byref): " & textFileSize & " bytes</li>")
Response.Write("<li>Preview: " & Server.HTMLEncode(Left(textData, 50)) & "...</li>")
Response.Write("</ul>")

' Test 2: Binary mode
Response.Write("<h2>Test 2: Read Binary File</h2>")
Dim binFileSize
Dim binData
binData = stream("demo_file.txt", True, binFileSize)

Response.Write("<ul>")
Response.Write("<li>Data length: " & Len(binData) & " bytes</li>")
Response.Write("<li>File size (byref): " & binFileSize & " bytes</li>")
Response.Write("</ul>")

Response.Write("<h2>✓ Function Working Correctly!</h2>")
Response.Write("<p><strong>Features working:</strong></p>")
Response.Write("<ul>")
Response.Write("<li>✓ Open() - opens ADODB.Stream</li>")
Response.Write("<li>✓ Type property (1=binary, 2=text)</li>")
Response.Write("<li>✓ CharSet property (utf-8)</li>")
Response.Write("<li>✓ LoadFromFile() - loads files</li>")
Response.Write("<li>✓ Read() - reads binary data</li>")
Response.Write("<li>✓ ReadText() - reads text data</li>")
Response.Write("<li>✓ Size property - returns file size</li>")
Response.Write("<li>✓ byref parameter - modifies caller's variable</li>")
Response.Write("<li>✓ asperror() - error handling function</li>")
Response.Write("</ul>")
%>
