<%
On Error Resume Next

Response.Write "<h1>RegExp Debug Test</h1>"

' Test 1: Create object
Response.Write "<h2>Test 1: Create Object</h2>"
Dim regex
Set regex = Server.CreateObject("G3REGEXP")

If Err.Number = 0 Then
    Response.Write "✓ Object created successfully<br>"
Else
    Response.Write "✗ Error: " & Err.Description & "<br>"
End If
Err.Clear

' Test 2: Call Test() directly
Response.Write "<h2>Test 2: Direct Test() Call</h2>"
Dim result
result = regex.Test("12345")
Response.Write "Test('12345') returned: " & CStr(result) & "<br>"
Response.Write "Type: " & TypeName(result) & "<br>"

' Test 3: Call Test() with pattern
Response.Write "<h2>Test 3: Test() with different patterns</h2>"
result = regex.Test("hello")
Response.Write "Test('hello'): " & CStr(result) & "<br>"

result = regex.Test("")
Response.Write "Test(''): " & CStr(result) & "<br>"

result = regex.Test("anything")
Response.Write "Test('anything'): " & CStr(result) & "<br>"

' Test 4: Try GetProperty
Response.Write "<h2>Test 4: GetProperty Access</h2>"
Dim pattern
pattern = regex.GetProperty("pattern")
Response.Write "GetProperty('pattern'): " & CStr(pattern) & "<br>"

' Test 5: Try SetProperty and then GetProperty
Response.Write "<h2>Test 5: SetProperty and GetProperty</h2>"
regex.SetProperty "pattern", "test"
If Err.Number = 0 Then
    Response.Write "✓ SetProperty succeeded<br>"
Else
    Response.Write "✗ SetProperty error: " & Err.Description & "<br>"
End If
Err.Clear

pattern = regex.GetProperty("pattern")
Response.Write "After SetProperty 'test', GetProperty returned: " & CStr(pattern) & "<br>"

' Test 6: Call Execute
Response.Write "<h2>Test 6: Execute() Method</h2>"
Dim matches
Set matches = regex.Execute("test")
If Err.Number = 0 Then
    Response.Write "✓ Execute succeeded<br>"
    If Not matches Is Nothing Then
        Response.Write "Matches object type: " & TypeName(matches) & "<br>"
        On Error Resume Next
        Response.Write "Matches.Count: " & matches.Count & "<br>"
        If Err.Number <> 0 Then
            Response.Write "Error accessing Count: " & Err.Description & "<br>"
        End If
    Else
        Response.Write "Matches is Nothing<br>"
    End If
Else
    Response.Write "✗ Execute error: " & Err.Description & "<br>"
End If
Err.Clear

' Test 7: Check object interface
Response.Write "<h2>Test 7: Object Interface Check</h2>"
Response.Write "Object type: " & TypeName(regex) & "<br>"

%>
Response.Write "Starting test<br>"

Set objRegExp = New RegExp
Response.Write "Created RegExp object<br>"

objRegExp.IgnoreCase = True
Response.Write "Set IgnoreCase<br>"

objRegExp.Pattern = "hello"
Response.Write "Set Pattern<br>"

Response.Write "Pattern: " & objRegExp.Pattern & "<br>"
Response.Write "IgnoreCase: " & objRegExp.IgnoreCase & "<br>"

Response.Write "Test completed!<br>"
%>
