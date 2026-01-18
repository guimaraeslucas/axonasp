<html>
<head>
    <title>G3REGEXP - Regular Expressions Test</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; }
        h1 { color: #0066cc; border-bottom: 3px solid #0066cc; padding-bottom: 10px; }
        h2 { color: #333; margin-top: 30px; border-left: 4px solid #0066cc; padding-left: 10px; }
        .test-section { margin: 20px 0; padding: 15px; background: #f9f9f9; border-left: 4px solid #0066cc; }
        .success { color: #28a745; font-weight: bold; }
        .error { color: #dc3545; font-weight: bold; }
        .info { color: #0066cc; font-weight: bold; }
        code { background: #f0f0f0; padding: 2px 5px; border-radius: 3px; font-family: 'Courier New'; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; background: white; }
        table td { padding: 8px; border: 1px solid #ddd; }
        table td:first-child { font-weight: bold; background: #f0f0f0; width: 300px; }
        .result { font-family: monospace; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3REGEXP - Regular Expression Engine Test Suite</h1>
        <p>Complete testing of RegExp object implementation using Go's regexp engine</p>

<%
On Error Resume Next

Dim regex, result, testStr, matches, match, i, count

' Test 1: Object Creation
Response.Write "<div class='test-section'>"
Response.Write "<h2>Test 1: Object Creation</h2>"
Response.Write "<table><tr><th>Test</th><th>Result</th></tr>"

Set regex = New RegExp
If Err.Number = 0 Then
    Response.Write "<tr><td>Set regex = New RegExp</td><td class='success'>SUCCESS</td></tr>"
Else
    Response.Write "<tr><td>Set regex = New RegExp</td><td class='error'>FAILED: " & Err.Description & "</td></tr>"
End If
Err.Clear

Set regex = Server.CreateObject("REGEXP")
If Err.Number = 0 Then
    Response.Write "<tr><td>Server.CreateObject('REGEXP') - alias</td><td class='success'>SUCCESS</td></tr>"
Else
    Response.Write "<tr><td>Server.CreateObject('REGEXP')</td><td class='error'>FAILED</td></tr>"
End If

Response.Write "</table></div>"
Err.Clear

' Test 2: Pattern property
Response.Write "<div class='test-section'>"
Response.Write "<h2>Test 2: Pattern Property (Direct Assignment)</h2>"
Response.Write "<table><tr><th>Test</th><th>Result</th></tr>"

Set regex = New RegExp
regex.Pattern = "\d+"
Response.Write "<tr><td>regex.Pattern = '\d+'</td><td class='result'>" & regex.Pattern & "</td></tr>"

regex.Pattern = "[a-z]+"
Response.Write "<tr><td>regex.Pattern = '[a-z]+'</td><td class='result'>" & regex.Pattern & "</td></tr>"

Response.Write "</table></div>"

' Test 3: Boolean properties
Response.Write "<div class='test-section'>"
Response.Write "<h2>Test 3: Boolean Properties</h2>"
Response.Write "<table><tr><th>Test</th><th>Result</th></tr>"

Set regex = New RegExp
regex.IgnoreCase = True
Response.Write "<tr><td>regex.IgnoreCase = True</td><td class='result'>" & regex.IgnoreCase & "</td></tr>"

regex.Global = True
Response.Write "<tr><td>regex.Global = True</td><td class='result'>" & regex.Global & "</td></tr>"

Response.Write "</table></div>"

' Test 4: Test() method
Response.Write "<div class='test-section'>"
Response.Write "<h2>Test 4: Test() Method (Pattern Matching)</h2>"
Response.Write "<table><tr><th>Test</th><th>Pattern</th><th>String</th><th>Result</th></tr>"

Set regex = New RegExp
regex.Pattern = "^[a-z]+$"
regex.IgnoreCase = False

testStr = "hello"
result = regex.Test(testStr)
Response.Write "<tr><td>Case sensitive match</td><td class='result'>^[a-z]+$</td><td class='result'>" & testStr & "</td>"
If result Then
    Response.Write "<td class='success'>MATCH</td>"
Else
    Response.Write "<td class='error'>NO MATCH</td>"
End If
Response.Write "</tr>"

testStr = "Hello"
result = regex.Test(testStr)
Response.Write "<tr><td>Case sensitive mismatch</td><td class='result'>^[a-z]+$</td><td class='result'>" & testStr & "</td>"
If result Then
    Response.Write "<td class='error'>UNEXPECTED MATCH</td>"
Else
    Response.Write "<td class='success'>CORRECTLY NO MATCH</td>"
End If
Response.Write "</tr>"

regex.IgnoreCase = True
testStr = "HELLO"
result = regex.Test(testStr)
Response.Write "<tr><td>Case insensitive match</td><td class='result'>^[a-z]+$ (IgnoreCase)</td><td class='result'>" & testStr & "</td>"
If result Then
    Response.Write "<td class='success'>MATCH</td>"
Else
    Response.Write "<td class='error'>NO MATCH</td>"
End If
Response.Write "</tr>"

regex.Pattern = "\d{3}-\d{4}"
testStr = "555-1234"
result = regex.Test(testStr)
Response.Write "<tr><td>Phone number pattern</td><td class='result'>\d{3}-\d{4}</td><td class='result'>" & testStr & "</td>"
If result Then
    Response.Write "<td class='success'>MATCH</td>"
Else
    Response.Write "<td class='error'>NO MATCH</td>"
End If
Response.Write "</tr>"

Response.Write "</table></div>"

' Test 5: Replace() method
Response.Write "<div class='test-section'>"
Response.Write "<h2>Test 5: Replace() Method</h2>"
Response.Write "<table><tr><th>Test</th><th>Original</th><th>Pattern</th><th>Replacement</th><th>Result</th></tr>"

Set regex = New RegExp

' Replace all numbers
regex.Pattern = "\d+"
regex.Global = True
testStr = "I have 10 apples and 25 oranges"
result = regex.Replace(testStr, "[NUM]")
Response.Write "<tr><td>Replace all numbers</td><td class='result'>" & testStr & "</td><td class='result'>\d+</td><td class='result'>[NUM]</td><td class='result'>" & result & "</td></tr>"

' Replace vowels
regex.Pattern = "[aeiou]"
regex.IgnoreCase = True
testStr = "Hello World"
result = regex.Replace(testStr, "*")
Response.Write "<tr><td>Replace vowels</td><td class='result'>" & testStr & "</td><td class='result'>[aeiou]</td><td class='result'>*</td><td class='result'>" & result & "</td></tr>"

' Replace first only
regex.Pattern = "\d+"
regex.Global = False
testStr = "first 123 second 456 third 789"
result = regex.Replace(testStr, "X")
Response.Write "<tr><td>Replace first match only</td><td class='result'>" & testStr & "</td><td class='result'>\d+</td><td class='result'>X</td><td class='result'>" & result & "</td></tr>"

Response.Write "</table></div>"

' Test 6: Execute() method
Response.Write "<div class='test-section'>"
Response.Write "<h2>Test 6: Execute() Method (Find Matches)</h2>"

Set regex = New RegExp
regex.Pattern = "\d+"
regex.Global = True

testStr = "The code is 2024, user id 12345, pin 999"
Set matches = regex.Execute(testStr)

If Err.Number = 0 Then
    Response.Write "<p><strong>Input:</strong> <code>" & testStr & "</code></p>"
    Response.Write "<p><strong>Pattern:</strong> <code>\d+</code> (Global=True)</p>"
    If Not matches Is Nothing Then
        count = matches.Count
        Response.Write "<p><strong>Matches found: " & count & "</strong></p>"
        If count > 0 Then
            Response.Write "<table><tr><th>#</th><th>Match Value</th><th>Index</th><th>Length</th></tr>"
            For i = 0 To count - 1
                Set match = matches.Item(i)
                Response.Write "<tr><td>" & (i+1) & "</td><td class='result'>" & match.Value & "</td><td>" & match.FirstIndex & "</td><td>" & match.Length & "</td></tr>"
            Next
            Response.Write "</table>"
        End If
    End If
Else
    Response.Write "<p class='error'>ERROR: " & Err.Description & "</p>"
End If
Err.Clear

Response.Write "</div>"

' Test 7: Execute with Global=False
Response.Write "<div class='test-section'>"
Response.Write "<h2>Test 7: Execute() with Global=False (First Match Only)</h2>"

Set regex = New RegExp
regex.Pattern = "[a-z]+"
regex.IgnoreCase = True
regex.Global = False

testStr = "Hello World JavaScript 2024"
Set matches = regex.Execute(testStr)

Response.Write "<p><strong>Input:</strong> <code>" & testStr & "</code></p>"
Response.Write "<p><strong>Pattern:</strong> <code>[a-z]+</code> (IgnoreCase=True, Global=False)</p>"

If matches.Count > 0 Then
    Response.Write "<p><strong>First match found:</strong></p>"
    Set match = matches.Item(0)
    Response.Write "<p>Value: <code>" & match.Value & "</code>, Index: " & match.FirstIndex & ", Length: " & match.Length & "</p>"
Else
    Response.Write "<p>No matches found</p>"
End If

Response.Write "</div>"

' Test 8: Common patterns
Response.Write "<div class='test-section'>"
Response.Write "<h2>Test 8: Common Validation Patterns</h2>"
Response.Write "<table><tr><th>Pattern Type</th><th>Test String</th><th>Valid</th></tr>"

Set regex = New RegExp

' Email pattern
regex.Pattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
testStr = "user@example.com"
result = regex.Test(testStr)
Response.Write "<tr><td>Email</td><td class='result'>" & testStr & "</td>"
If result Then
    Response.Write "<td class='success'>VALID</td>"
Else
    Response.Write "<td class='error'>INVALID</td>"
End If
Response.Write "</tr>"

' URL pattern
regex.Pattern = "^https?://"
testStr = "https://www.example.com"
result = regex.Test(testStr)
Response.Write "<tr><td>URL (http/https)</td><td class='result'>" & testStr & "</td>"
If result Then
    Response.Write "<td class='success'>VALID</td>"
Else
    Response.Write "<td class='error'>INVALID</td>"
End If
Response.Write "</tr>"

' IP address
regex.Pattern = "^[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}$"
testStr = "192.168.1.1"
result = regex.Test(testStr)
Response.Write "<tr><td>IP Address</td><td class='result'>" & testStr & "</td>"
If result Then
    Response.Write "<td class='success'>VALID</td>"
Else
    Response.Write "<td class='error'>INVALID</td>"
End If
Response.Write "</tr>"

Response.Write "</table></div>"

' Footer
Response.Write "<footer style='margin-top: 40px; padding-top: 20px; border-top: 2px solid #0066cc; text-align: center;'>"
Response.Write "<p><strong>G3 AxonASP</strong> - RegExp Complete Implementation</p>"
Response.Write "<p>Powered by Go's regexp engine with full VBScript compatibility</p>"
Response.Write "</footer>"
%>

    </div>
</body>
</html>
