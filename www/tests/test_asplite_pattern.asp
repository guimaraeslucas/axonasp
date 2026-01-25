<%@ language="VBScript" %>
<% Option Explicit %>

<h1>ASPLite Pattern Test</h1>
<pre>

<% 
' Test the actual aspLite pattern: colon + End If on same line
' Pattern: if instr(...) > 0 then ... : end if

Dim blockURL, testStr, result

' Simple test
testStr = "hello"
result = ""
If InStr(testStr, "l") > 0 Then result = "found" : End If
Response.Write "Test 1 (InStr > 0): " & result & "<br />"

' Test 2: Negative case
testStr = "world"
result = "notfound"
If InStr(testStr, "x") > 0 Then result = "found" : End If
Response.Write "Test 2 (InStr = 0): " & result & "<br />"

' Test 3: Multiple statements before End If
Dim value
value = 5
If value > 3 Then value = value + 1 : value = value * 2 : End If
Response.Write "Test 3 (Multiple statements): " & value & "<br />"

' Test 4: Nested Case with colon
Dim numTest, caseResult
numTest = 2
caseResult = "unknown"
Select Case numTest
    Case 1
        caseResult = "one"
    Case 2 : caseResult = "two" : caseResult = caseResult & "-processed"
    Case Else
        caseResult = "other"
End Select
Response.Write "Test 4 (Case with colon): " & caseResult & "<br />"

%>

</pre>
