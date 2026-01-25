<%@ Language=VBScript %>
<%
' Test File: Nil/Empty Path Validation
' Tests that ADODB.Stream and FSO properly handle nil/empty paths
Response.ContentType = "text/html"

Response.Write("<html><head><title>Nil Path Validation Test</title></head><body>")
Response.Write("<h1>Nil/Empty Path Validation Test</h1>")

' ======= TEST 1: ADODB.Stream with empty path =======
Response.Write("<h2>Test 1: ADODB.Stream LoadFromFile with empty path</h2>")
Dim objStream
Set objStream = server.CreateObject("ADODB.Stream")
objStream.Open
objStream.type = 2

Response.Write("<p>Attempting LoadFromFile with empty string...</p>")
objStream.LoadFromFile("")
Response.Write("<p style='color:green'>✓ No crash - check console for proper error message</p>")

objStream.Close
Set objStream = Nothing

' ======= TEST 2: ADODB.Stream with valid path =======
Response.Write("<h2>Test 2: ADODB.Stream with valid file</h2>")
Set objStream = server.CreateObject("ADODB.Stream")
objStream.Open
objStream.type = 2
objStream.LoadFromFile(server.MapPath("demo_file.txt"))
Response.Write("<p style='color:green'>✓ Valid file loaded successfully (Size: " & objStream.Size & " bytes)</p>")
objStream.Close
Set objStream = Nothing

' ======= TEST 3: G3FILES with empty path =======
Response.Write("<h2>Test 3: G3FILES with empty path</h2>")
Dim files
Set files = server.CreateObject("G3FILES")

Response.Write("<p>Attempting Exists with empty string...</p>")
Dim result
result = files.Exists("")
Response.Write("<p style='color:green'>✓ No crash - check console for proper error message (returned: " & result & ")</p>")

Set files = Nothing

' ======= TEST 4: G3FILES with valid path =======
Response.Write("<h2>Test 4: G3FILES with valid file</h2>")
Set files = server.CreateObject("G3FILES")
result = files.Exists("demo_file.txt")
Response.Write("<p style='color:green'>✓ Valid file checked (exists: " & result & ")</p>")
Set files = Nothing

' ======= SUMMARY =======
Response.Write("<hr>")
Response.Write("<h2>✓ Test Complete</h2>")
Response.Write("<p>All methods handled nil/empty paths without crashing.</p>")
Response.Write("<p><strong>Check the server console logs for:</strong></p>")
Response.Write("<ul>")
Response.Write("<li>✓ 'Error: LoadFromFile received empty or nil filename' message</li>")
Response.Write("<li>✓ 'Error: G3FILES received empty or nil path' message</li>")
Response.Write("<li>✓ NO '<nil>' in security warnings</li>")
Response.Write("</ul>")

Response.Write("</body></html>")
%>
