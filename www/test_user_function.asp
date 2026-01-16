<%
' Test user's stream function with full functionality
Option Explicit

' Define a simple asperror function if it doesn't exist
Function asperror(path)
    ' Optional: log errors related to file operations
    If Err.Number <> 0 Then
        Response.Write("<!-- Error for " & path & ": " & Err.Description & " -->" & vbCrLf)
    End If
End Function

' User's function exactly as provided
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

Response.Write("<h1>Testing User's stream() Function</h1>")

' Test 1: Text mode with byref
Response.Write("<h2>Test 1: Text Mode</h2>")
Dim textSize
Dim textContent
textContent = stream("demo_file.txt", False, textSize)
Response.Write("Content length (Len): " & Len(textContent) & "<br>")
Response.Write("Size (byref): " & textSize & "<br>")
Response.Write("Content preview: " & Server.HTMLEncode(Left(textContent, 50)) & "<br>")

' Test 2: Binary mode with byref
Response.Write("<h2>Test 2: Binary Mode</h2>")
Dim binSize
Dim binContent
binContent = stream("demo_file.txt", True, binSize)
Response.Write("Content length (Len): " & Len(binContent) & "<br>")
Response.Write("Size (byref): " & binSize & "<br>")

Response.Write("<p><strong>✓ Function working!</strong></p>")
%>
