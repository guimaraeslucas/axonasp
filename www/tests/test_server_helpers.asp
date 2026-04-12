<%
' Test Server Helper Methods
' MapPath, URLEncode, HTMLEncode, GetLastError, IsClientConnected

Response.Write("<h1>Server Helper Methods Test</h1>")
Response.Write("<hr>")

' Test 1: MapPath
Response.Write("<h2>1. Server.MapPath Test</h2>")
Response.Write("<p>MapPath('/test.txt'): " & Server.MapPath("/test.txt") & "</p>")
Response.Write("<p>MapPath('page.asp'): " & Server.MapPath("page.asp") & "</p>")

' Test 2: URLEncode
Response.Write("<h2>2. Server.URLEncode Test</h2>")
Dim testString
testString = "Hello World & Friends"
Response.Write("<p>Original: " & testString & "</p>")
Response.Write("<p>Encoded: " & Server.URLEncode(testString) & "</p>")

' Test 3: HTMLEncode
Response.Write("<h2>3. Server.HTMLEncode Test</h2>")
Dim htmlTest
htmlTest = "<script>alert('XSS')</script>"
Response.Write("<p>Original: " & htmlTest & "</p>")
Response.Write("<p>Encoded: " & Server.HTMLEncode(htmlTest) & "</p>")

' Test 4: IsClientConnected
Response.Write("<h2>4. Server.IsClientConnected Test</h2>")
Response.Write("<p>Client Connected: " & Server.IsClientConnected() & "</p>")

' Test 5: GetLastError
Response.Write("<h2>5. Server.GetLastError Test</h2>")
Dim lastErr
Set lastErr = Server.GetLastError()
If IsNull(lastErr) Then
    Response.Write("<p>No error occurred (expected)</p>")
Else
    Response.Write("<p>Error Number: " & lastErr & "</p>")
End If

Response.Write("<hr>")
Response.Write("<p>All tests completed successfully!</p>")
%>
