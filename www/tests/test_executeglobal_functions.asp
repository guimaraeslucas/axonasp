<%
Response.Write "<h1>ExecuteGlobal Function Test</h1>"

' Test 1: Function without parameters
Response.Write "<h2>Test 1: Simple function</h2>"
Dim code1
code1 = "Function Test1()" & vbCrLf & _
        "Test1 = ""result1""" & vbCrLf & _
        "End Function"
Response.Write "<p>Code: <pre>" & Server.HTMLEncode(code1) & "</pre></p>"
ExecuteGlobal code1
Response.Write "<p>Result: Test1() = " & Test1() & "</p>"

' Test 2: Function with parameter (version 1: simple return)
Response.Write "<h2>Test 2: Function with parameter (simple return)</h2>"
Dim code2
code2 = "Function Test2(x)" & vbCrLf & _
        "Test2 = x" & vbCrLf & _
        "End Function"
Response.Write "<p>Code: <pre>" & Server.HTMLEncode(code2) & "</pre></p>"
ExecuteGlobal code2
Response.Write "<p>Result: Test2(""hello"") = " & Test2("hello") & "</p>"

' Test 3: Function with parameter and concatenation
Response.Write "<h2>Test 3: Function with concatenation</h2>"
Dim code3, q
q = Chr(34)
code3 = "Function Test3(x)" & vbCrLf & _
        "Test3 = x & " & q & " processed" & q & vbCrLf & _
        "End Function"
Response.Write "<p>Code: <pre>" & Server.HTMLEncode(code3) & "</pre></p>"
ExecuteGlobal code3
Response.Write "<p>Result: Test3(""hello"") = " & Test3("hello") & "</p>"

Response.Write "<p>All tests complete</p>"
%>
