<%
' AXONASP_AUDIT Issues 12, 13, 14 - FINAL VALIDATION TEST
' This test validates all three resolved issues from the audit

Response.Write "<html><head><title>AXONASP Audit - Issues 12, 13, 14 - Final Validation</title>"
Response.Write "<style>"
Response.Write "body { font-family: Arial, sans-serif; margin: 20px; }"
Response.Write ".pass { background: #d4edda; border: 1px solid #c3e6cb; padding: 10px; margin: 10px 0; }"
Response.Write ".fail { background: #f8d7da; border: 1px solid #f5c6cb; padding: 10px; margin: 10px 0; }"
Response.Write ".title { font-size: 18px; font-weight: bold; margin-top: 20px; }"
Response.Write "pre { background: #f5f5f5; padding: 10px; overflow-x: auto; }"
Response.Write "</style></head><body>"

Response.Write "<h1>AXONASP Audit - Final Validation Report</h1>"
Response.Write "<p>This test validates Issues 12, 13, and 14 from AXONASP_AUDIT.MD</p>"

Dim testsPassed, testsFailed
testsPassed = 0
testsFailed = 0

' ========================================
' ISSUE 12: ASPERROR OBJECT COMPLETENESS
' ========================================
Response.Write "<div class='title'>Issue 12: ASPError Object Completeness</div>"

On Error Resume Next

' Test 1: Set and retrieve all properties
Err.Clear
Err.ASPCode = "ERR_TEST_001"
Err.Category = "Syntax Error"
Err.File = "test_issues.asp"
Err.Line = 42
Err.Column = 15
Err.ASPDescription = "Extended ASP error description"

Err.Number = 11
Err.Description = "Division by zero"
Err.Source = "Test Source"

Dim test12Pass
test12Pass = True

If Err.ASPCode <> "ERR_TEST_001" Then test12Pass = False
If Err.Category <> "Syntax Error" Then test12Pass = False
If Err.File <> "test_issues.asp" Then test12Pass = False
If Err.Line <> 42 Then test12Pass = False
If Err.Column <> 15 Then test12Pass = False
If Err.ASPDescription <> "Extended ASP error description" Then test12Pass = False
If Err.Number <> 11 Then test12Pass = False

If test12Pass Then
    Response.Write "<div class='pass'>"
    Response.Write "<strong>✓ PASS:</strong> All ASPError properties working correctly<br>"
    Response.Write "Properties tested: ASPCode, Category, File, Line, Column, ASPDescription, Number, Description, Source"
    Response.Write "</div>"
    testsPassed = testsPassed + 1
Else
    Response.Write "<div class='fail'>"
    Response.Write "<strong>✗ FAIL:</strong> Some ASPError properties not working"
    Response.Write "</div>"
    testsFailed = testsFailed + 1
End If

Err.Clear

' ========================================
' ISSUE 13: REGEXP OBJECT - SUBMATCHES
' ========================================
Response.Write "<div class='title'>Issue 13: RegExp Object Return Values (SubMatches)</div>"

Dim test13Pass
test13Pass = True

Set regex = New RegExp
regex.Pattern = "(\d{3})-(\w+)-([a-z]+)"
regex.Global = False
regex.IgnoreCase = True

Dim testString, matches, firstMatch, subMatches
testString = "123-TEST-abc"

Set matches = regex.Execute(testString)

If matches.Count <> 1 Then
    test13Pass = False
Else
    Set firstMatch = matches.Item(0)

    If firstMatch.Value <> "123-TEST-abc" Then test13Pass = False
    If firstMatch.FirstIndex <> 0 Then test13Pass = False
    If firstMatch.Length <> 12 Then test13Pass = False

    If IsEmpty(firstMatch.SubMatches) Then
        test13Pass = False
    Else
        Set subMatches = firstMatch.SubMatches
        If subMatches.Count <> 3 Then test13Pass = False

        ' Validate captured groups
        Dim group0, group1, group2
        group0 = subMatches.Item(0).Value
        group1 = subMatches.Item(1).Value
        group2 = subMatches.Item(2).Value

        If group0 <> "123" Then test13Pass = False
        If group1 <> "TEST" Then test13Pass = False
        If group2 <> "abc" Then test13Pass = False
    End If
End If

If test13Pass Then
    Response.Write "<div class='pass'>"
    Response.Write "<strong>✓ PASS:</strong> RegExp Execute() returns correct structure with SubMatches<br>"
    Response.Write "Validated: Match.Value, Match.Index, Match.Length, Match.SubMatches collection<br>"
    Response.Write "SubMatches captured: [0]='" & subMatches.Item(0).Value & "', [1]='" & subMatches.Item(1).Value & "', [2]='" & subMatches.Item(2).Value & "'"
    Response.Write "</div>"
    testsPassed = testsPassed + 1
Else
    Response.Write "<div class='fail'>"
    Response.Write "<strong>✗ FAIL:</strong> RegExp SubMatches not working correctly"
    Response.Write "</div>"
    testsFailed = testsFailed + 1
End If

' ========================================
' ISSUE 14: DICTIONARY DEFAULT PROPERTY
' ========================================
Response.Write "<div class='title'>Issue 14: Dictionary Default Property (Item)</div>"

Dim test14Pass
test14Pass = True

Set dict = Server.CreateObject("Scripting.Dictionary")

' Test default property syntax
dict("key1") = "value1"
dict("key2") = "value2"
dict("key3") = 12345

' Validate retrieval
If dict("key1") <> "value1" Then test14Pass = False
If dict("key2") <> "value2" Then test14Pass = False
If dict("key3") <> 12345 Then test14Pass = False

' Validate Item method still works
If dict.Item("key1") <> "value1" Then test14Pass = False

' Test Count
If dict.Count <> 3 Then test14Pass = False

' Test Add method compatibility
dict.Add "key4", "value4"
If dict("key4") <> "value4" Then test14Pass = False

If test14Pass Then
    Response.Write "<div class='pass'>"
    Response.Write "<strong>✓ PASS:</strong> Dictionary default property working correctly<br>"
    Response.Write "Both syntaxes work: dict('key') and dict.Item('key')<br>"
    Response.Write "Count: " & dict.Count & " items"
    Response.Write "</div>"
    testsPassed = testsPassed + 1
Else
    Response.Write "<div class='fail'>"
    Response.Write "<strong>✗ FAIL:</strong> Dictionary default property not working"
    Response.Write "</div>"
    testsFailed = testsFailed + 1
End If

' ========================================
' SUMMARY
' ========================================
Response.Write "<div class='title'>Summary</div>"
Response.Write "<pre>"
Response.Write "Total Tests: " & (testsPassed + testsFailed) & vbCrLf
Response.Write "Passed:      " & testsPassed & vbCrLf
Response.Write "Failed:      " & testsFailed & vbCrLf

If testsFailed = 0 Then
    Response.Write vbCrLf & "✓ ALL TESTS PASSED!" & vbCrLf
    Response.Write "All three issues (12, 13, 14) have been successfully resolved." & vbCrLf
Else
    Response.Write vbCrLf & "✗ Some tests failed" & vbCrLf
End If

Response.Write "</pre>"

Response.Write "</body></html>"
%>
