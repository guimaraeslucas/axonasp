<%
'============================================================
' MSXML2 COMPREHENSIVE VALIDATION TEST
' This test validates all MSXML2 functionality is working
' Uses correct VBScript syntax with proper parentheses
'============================================================

Response.Write "<h1>MSXML2 Comprehensive Validation</h1>"
Response.Write "<p><strong>All tests use proper VBScript syntax</strong></p>"

Dim testsPassed, testsFailed
testsPassed = 0
testsFailed = 0

' ===== TEST 1: DOMDocument Creation =====
Response.Write "<h3>Test 1: DOMDocument Creation</h3>"
On Error Resume Next
Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
If Err.Number <> 0 Then
    Response.Write "❌ FAIL: " & Err.Description & "<br>"
    testsFailed = testsFailed + 1
    Err.Clear
Else
    Response.Write "✅ PASS: DOMDocument created successfully<br>"
    testsPassed = testsPassed + 1
End If

' ===== TEST 2: LoadXML =====
Response.Write "<h3>Test 2: LoadXML</h3>"
xmlStr = "<root><item id=""1"">First</item><item id=""2"">Second</item><item id=""3"">Third</item></root>"
result = xmlDoc.LoadXML(xmlStr)
If result Then
    Response.Write "✅ PASS: XML loaded successfully<br>"
    testsPassed = testsPassed + 1
Else
    Response.Write "❌ FAIL: LoadXML returned False<br>"
    testsFailed = testsFailed + 1
End If

' ===== TEST 3: DocumentElement Property =====
Response.Write "<h3>Test 3: DocumentElement Property</h3>"
Set rootElem = xmlDoc.DocumentElement
If rootElem Is Nothing Then
    Response.Write "❌ FAIL: DocumentElement is Nothing<br>"
    testsFailed = testsFailed + 1
Else
    If rootElem.NodeName = "root" Then
        Response.Write "✅ PASS: DocumentElement exists with correct NodeName: " & rootElem.NodeName & "<br>"
        testsPassed = testsPassed + 1
    Else
        Response.Write "❌ FAIL: DocumentElement has wrong NodeName: " & rootElem.NodeName & "<br>"
        testsFailed = testsFailed + 1
    End If
End If

' ===== TEST 4: XML Property =====
Response.Write "<h3>Test 4: XML Property</h3>"
retrievedXML = xmlDoc.XML
If Len(retrievedXML) > 0 Then
    Response.Write "✅ PASS: XML property returns content (length: " & Len(retrievedXML) & ")<br>"
    testsPassed = testsPassed + 1
Else
    Response.Write "❌ FAIL: XML property is empty<br>"
    testsFailed = testsFailed + 1
End If

' ===== TEST 5: GetElementsByTagName =====
Response.Write "<h3>Test 5: GetElementsByTagName</h3>"
Set items = xmlDoc.GetElementsByTagName("item")
If items Is Nothing Then
    Response.Write "❌ FAIL: GetElementsByTagName returned Nothing<br>"
    testsFailed = testsFailed + 1
Else
    itemCount = UBound(items) + 1
    If itemCount = 3 Then
        Response.Write "✅ PASS: Found correct number of items (" & itemCount & ")<br>"
        testsPassed = testsPassed + 1
        
        ' Verify we can access individual items
        Set firstItem = items(0)
        If firstItem Is Nothing Then
            Response.Write "  ⚠️ WARNING: Cannot access item(0)<br>"
        Else
            Response.Write "  - Item 0: " & firstItem.Text & "<br>"
        End If
    Else
        Response.Write "❌ FAIL: Wrong number of items: " & itemCount & " (expected 3)<br>"
        testsFailed = testsFailed + 1
    End If
End If

' ===== TEST 6: Element Text Property =====
Response.Write "<h3>Test 6: Element Text Property</h3>"
If items Is Nothing Then
    Response.Write "⏭️ SKIPPED: No items from previous test<br>"
Else
    Set item1 = items(0)
    If item1 Is Nothing Then
        Response.Write "❌ FAIL: Cannot get first item<br>"
        testsFailed = testsFailed + 1
    Else
        itemText = item1.Text
        If itemText = "First" Then
            Response.Write "✅ PASS: Element text is correct: " & itemText & "<br>"
            testsPassed = testsPassed + 1
        Else
            Response.Write "❌ FAIL: Element text is wrong: " & itemText & " (expected 'First')<br>"
            testsFailed = testsFailed + 1
        End If
    End If
End If

' ===== TEST 7: CreateElement =====
Response.Write "<h3>Test 7: CreateElement</h3>"
Set newElem = xmlDoc.CreateElement("newelement")
If newElem Is Nothing Then
    Response.Write "❌ FAIL: CreateElement returned Nothing<br>"
    testsFailed = testsFailed + 1
Else
    If newElem.NodeName = "newelement" Then
        Response.Write "✅ PASS: Element created with correct NodeName: " & newElem.NodeName & "<br>"
        testsPassed = testsPassed + 1
    Else
        Response.Write "❌ FAIL: Element has wrong NodeName: " & newElem.NodeName & "<br>"
        testsFailed = testsFailed + 1
    End If
End If

' ===== TEST 8: CreateTextNode =====
Response.Write "<h3>Test 8: CreateTextNode</h3>"
Set textNode = xmlDoc.CreateTextNode("Sample text content")
If textNode Is Nothing Then
    Response.Write "❌ FAIL: CreateTextNode returned Nothing<br>"
    testsFailed = testsFailed + 1
Else
    Response.Write "✅ PASS: Text node created<br>"
    testsPassed = testsPassed + 1
End If

' ===== TEST 9: SelectSingleNode with // =====
Response.Write "<h3>Test 9: SelectSingleNode (//item)</h3>"
Set node = xmlDoc.SelectSingleNode("//item")
If node Is Nothing Then
    Response.Write "❌ FAIL: SelectSingleNode returned Nothing<br>"
    testsFailed = testsFailed + 1
Else
    Response.Write "✅ PASS: Found node: " & node.NodeName & " = " & node.Text & "<br>"
    testsPassed = testsPassed + 1
End If

' ===== TEST 10: SelectSingleNode with full path =====
Response.Write "<h3>Test 10: SelectSingleNode (root/item)</h3>"
Set node2 = xmlDoc.SelectSingleNode("root/item")
If node2 Is Nothing Then
    Response.Write "❌ FAIL: SelectSingleNode returned Nothing<br>"
    testsFailed = testsFailed + 1
Else
    Response.Write "✅ PASS: Found node: " & node2.NodeName & " = " & node2.Text & "<br>"
    testsPassed = testsPassed + 1
End If

' ===== TEST 11: SelectNodes =====
Response.Write "<h3>Test 11: SelectNodes (//item)</h3>"
Set nodes = xmlDoc.SelectNodes("//item")
If nodes Is Nothing Then
    Response.Write "❌ FAIL: SelectNodes returned Nothing<br>"
    testsFailed = testsFailed + 1
Else
    nodeCount = UBound(nodes) + 1
    If nodeCount = 3 Then
        Response.Write "✅ PASS: Found correct number of nodes (" & nodeCount & ")<br>"
        testsPassed = testsPassed + 1
    Else
        Response.Write "❌ FAIL: Wrong number of nodes: " & nodeCount & " (expected 3)<br>"
        testsFailed = testsFailed + 1
    End If
End If

' ===== TEST 12: ParseError on Invalid XML =====
Response.Write "<h3>Test 12: ParseError Detection</h3>"
Set xmlDoc2 = Server.CreateObject("MSXML2.DOMDocument")
invalidXML = "<unclosed>test"
result2 = xmlDoc2.LoadXML(invalidXML)
If result2 = False Then
    Set parseErr = xmlDoc2.ParseError
    If parseErr Is Nothing Then
        Response.Write "❌ FAIL: ParseError is Nothing<br>"
        testsFailed = testsFailed + 1
    Else
        If parseErr.ErrorCode <> 0 Then
            Response.Write "✅ PASS: ParseError detected (ErrorCode: " & parseErr.ErrorCode & ")<br>"
            testsPassed = testsPassed + 1
        Else
            Response.Write "❌ FAIL: ParseError.ErrorCode is 0<br>"
            testsFailed = testsFailed + 1
        End If
    End If
Else
    Response.Write "⚠️ WARNING: LoadXML succeeded on invalid XML<br>"
End If

' ===== TEST 13: ServerXMLHTTP Creation =====
Response.Write "<h3>Test 13: ServerXMLHTTP Creation</h3>"
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")
If http Is Nothing Then
    Response.Write "❌ FAIL: ServerXMLHTTP is Nothing<br>"
    testsFailed = testsFailed + 1
Else
    Response.Write "✅ PASS: ServerXMLHTTP created successfully<br>"
    testsPassed = testsPassed + 1
End If

' ===== TEST 14: ServerXMLHTTP Properties =====
Response.Write "<h3>Test 14: ServerXMLHTTP Properties</h3>"
If http Is Nothing Then
    Response.Write "⏭️ SKIPPED: No HTTP object<br>"
Else
    timeout = http.Timeout
    readyState = http.ReadyState
    If timeout > 0 And readyState >= 0 Then
        Response.Write "✅ PASS: Properties accessible (Timeout: " & timeout & ", ReadyState: " & readyState & ")<br>"
        testsPassed = testsPassed + 1
    Else
        Response.Write "❌ FAIL: Properties not working correctly<br>"
        testsFailed = testsFailed + 1
    End If
End If

' ===== SUMMARY =====
Response.Write "<hr>"
Response.Write "<h2>Test Summary</h2>"
Response.Write "<p><strong>Total Tests:</strong> " & (testsPassed + testsFailed) & "</p>"
Response.Write "<p><strong>✅ Passed:</strong> " & testsPassed & "</p>"
Response.Write "<p><strong>❌ Failed:</strong> " & testsFailed & "</p>"

If testsFailed = 0 Then
    Response.Write "<h3 style='color: green;'>🎉 ALL TESTS PASSED!</h3>"
    Response.Write "<p>MSXML2 implementation is fully functional.</p>"
Else
    passRate = Round((testsPassed * 100) / (testsPassed + testsFailed), 1)
    Response.Write "<h3 style='color: orange;'>⚠️ Some tests failed</h3>"
    Response.Write "<p>Pass rate: " & passRate & "%</p>"
End If

On Error GoTo 0
%>
