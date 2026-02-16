<%@ Language="VBScript" CodePage=65001 %>
<!-- Test private function execution with byref -->
<% 
    ' Enable debugging
    debug_asp_code = "TRUE"

    ' Define private function
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

    ' Test the function
    Dim textData, textFileSize
    textFileSize = 0
    
    Response.Write("<h2>Testing Private Function stream()</h2>")
    Response.Write("<hr />")
    
    Response.Write("<p><strong>Test 1: Read text file</strong></p>")
    textData = stream("demo_file.txt", False, textFileSize)
    Response.Write("<p>File content length: " & Len(textData) & " characters</p>")
    Response.Write("<p>File size (byref): " & textFileSize & " bytes</p>")
    Response.Write("<p>Content preview: " & server.HTMLEncode(Left(textData, 100)) & "...</p>")
    Response.Write("<hr />")

    ' Show full content
    Response.Write("<h3>Full File Content:</h3>")
    Response.Write("<pre>" & server.HTMLEncode(textData) & "</pre>")
%>
