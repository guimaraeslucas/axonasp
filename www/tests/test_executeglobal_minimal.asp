<%
Response.Write "<h1>Minimal ExecuteGlobal Test</h1>"

' Test 1: Simple variable
Response.Write "<h2>Test 1: Simple variable</h2>"
Dim code1
code1 = "Dim testVar" & vbCrLf & "testVar = 42"
Response.Write "<p>Code: <pre>" & Server.HTMLEncode(code1) & "</pre></p>"
ExecuteGlobal code1
Response.Write "<p>Result: testVar = " & testVar & "</p>"

' Test 2: Simple concatenation
Response.Write "<h2>Test 2: Simple concatenation</h2>"
Dim code2
code2 = "Dim result2" & vbCrLf & "result2 = ""hello"" & "" world"""
Response.Write "<p>Code: <pre>" & Server.HTMLEncode(code2) & "</pre></p>"
ExecuteGlobal code2
Response.Write "<p>Result: result2 = " & result2 & "</p>"

' Test 3: Concatenation with variable
Response.Write "<h2>Test 3: Concatenation with variable</h2>"
Dim code3
code3 = "Dim input3, result3" & vbCrLf & _
        "input3 = ""hello""" & vbCrLf & _
        "result3 = input3 & "" world"""
Response.Write "<p>Code: <pre>" & Server.HTMLEncode(code3) & "</pre></p>"
ExecuteGlobal code3
Response.Write "<p>Result: result3 = " & result3 & "</p>"

Response.Write "<p>All tests complete</p>"
%>
