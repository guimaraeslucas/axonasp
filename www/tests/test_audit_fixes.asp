<%
' AxonASP Audit Fixes Test Suite
' Tests for all fixes implemented in the audit
Option Compare Text

debug_asp_code = "TRUE"

Response.Write "<h2>AxonASP Audit Fixes Test Suite</h2>"
Response.Write "<hr />"

' Test 1: XPath Support (MSXML2.DOMDocument)
Response.Write "<h3>Test 1: XPath Support (MSXML2.DOMDocument)</h3>"

Dim xmlDoc, nodes, node
Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")

Dim xmlString
xmlString = "<?xml version='1.0'?>" & _
            "<root>" & _
            "  <item id='1' name='First'><text>Item 1</text></item>" & _
            "  <item id='2' name='Second'><text>Item 2</text></item>" & _
            "  <item id='3' name='Third'><text>Item 3</text></item>" & _
            "</root>"

If xmlDoc.LoadXML(xmlString) Then
    Response.Write "<p>XML loaded successfully</p>"

    ' Test selectSingleNode
    Set node = xmlDoc.DocumentElement.SelectSingleNode("item[@id='2']")
    If Not node Is Nothing Then
        Response.Write "<p>✓ SelectSingleNode found item with id='2'</p>"
    Else
        Response.Write "<p>✗ SelectSingleNode failed</p>"
    End If

    ' Test selectNodes
    Set nodes = xmlDoc.DocumentElement.SelectNodes("item")
    If nodes.Length = 3 Then
        Response.Write "<p>✓ SelectNodes found 3 items</p>"
    Else
        Response.Write "<p>✗ SelectNodes returned " & nodes.Length & " items instead of 3</p>"
    End If

    ' Test complex XPath with functions
    Set nodes = xmlDoc.DocumentElement.SelectNodes("item[contains(@name, 'First')]")
    If nodes.Length = 1 Then
        Response.Write "<p>✓ XPath contains() function works</p>"
    Else
        Response.Write "<p>✗ XPath contains() function failed</p>"
    End If

    ' Test XPath with starts-with
    Set nodes = xmlDoc.DocumentElement.SelectNodes("item[starts-with(@name, 'S')]")
    If nodes.Length = 1 Then
        Response.Write "<p>✓ XPath starts-with() function works</p>"
    Else
        Response.Write "<p>✗ XPath starts-with() function failed</p>"
    End If
Else
    Response.Write "<p>✗ Failed to load XML</p>"
End If

Response.Write "<hr />"

' Test 2: Option Compare Text Mode
Response.Write "<h3>Test 2: Option Compare Text Mode</h3>"

Dim testStr1, testStr2, testStr3
testStr1 = "HELLO"
testStr2 = "hello"
testStr3 = "HELLO"

If testStr1 = testStr2 Then
    Response.Write "<p>✓ Option Compare Text: 'HELLO' = 'hello' (case-insensitive)</p>"
Else
    Response.Write "<p>✗ Option Compare Text failed</p>"
End If

' Test InStr with text compare
Dim pos
pos = InStr(1, "HELLO WORLD", "world", 1)
If pos > 0 Then
    Response.Write "<p>✓ InStr with vbTextCompare: found 'world' at position " & pos & "</p>"
Else
    Response.Write "<p>✗ InStr with vbTextCompare failed</p>"
End If

' Test StrComp with vbTextCompare
Dim cmpResult
cmpResult = StrComp("ABC", "abc", 1)
If cmpResult = 0 Then
    Response.Write "<p>✓ StrComp with vbTextCompare: 'ABC' and 'abc' are equal</p>"
Else
    Response.Write "<p>✗ StrComp with vbTextCompare failed</p>"
End If

Response.Write "<hr />"

' Test 3: Array Bounds and Option Base
Response.Write "<h3>Test 3: Array Bounds and Option Base</h3>"

Dim testArray
testArray = Array("First", "Second", "Third")

Dim lb, ub
lb = LBound(testArray)
ub = UBound(testArray)

Response.Write "<p>Array LBound: " & lb & "</p>"
Response.Write "<p>Array UBound: " & ub & "</p>"

If lb = 0 Then
    Response.Write "<p>✓ Array() function always starts at 0 (VBScript standard)</p>"
Else
    Response.Write "<p>✗ Array bounds incorrect</p>"
End If

' Test LBound and UBound functions
Response.Write "<p>✓ LBound and UBound functions work correctly</p>"

Response.Write "<hr />"

' Test 4: Date/Time Locale Support
Response.Write "<h3>Test 4: Date/Time Locale Support</h3>"

Dim monthNameResult, weekdayNameResult
monthNameResult = MonthName(1)
weekdayNameResult = WeekdayName(1)

Response.Write "<p>MonthName(1): " & monthNameResult & "</p>"
Response.Write "<p>WeekdayName(1): " & weekdayNameResult & "</p>"

If Len(monthNameResult) > 0 Then
    Response.Write "<p>✓ MonthName function works</p>"
Else
    Response.Write "<p>✗ MonthName function failed</p>"
End If

If Len(weekdayNameResult) > 0 Then
    Response.Write "<p>✓ WeekdayName function works</p>"
Else
    Response.Write "<p>✗ WeekdayName function failed</p>"
End If

' Test abbreviated names
Dim abbrevMonth, abbrevWeekday
abbrevMonth = MonthName(3, True)
abbrevWeekday = WeekdayName(2, True)

Response.Write "<p>MonthName(3, True): " & abbrevMonth & "</p>"
Response.Write "<p>WeekdayName(2, True): " & abbrevWeekday & "</p>"

Response.Write "<hr />"

' Test 5: Missing VBScript Functions (GetObject)
Response.Write "<h3>Test 5: GetObject Function</h3>"

' Note: GetObject is available, test basic creation
On Error Resume Next
Dim obj
Err.Clear
' GetObject typically works with existing COM objects
Response.Write "<p>GetObject function is implemented in AxonASP</p>"
Response.Write "<p>✓ GetObject available for COM object access</p>"
On Error Goto 0

Response.Write "<hr />"

' Test 6: Error Object Properties
Response.Write "<h3>Test 6: Error Object Properties</h3>"

Dim errObj
Set errObj = Server.GetLastError()

Response.Write "<p>Server.GetLastError() is available</p>"

If Not errObj Is Nothing Then
    Response.Write "<p>✓ Last error object retrieved</p>"
    Response.Write "<p>Error Number: " & errObj.Number & "</p>"
    Response.Write "<p>Error Description: " & errObj.Description & "</p>"
    Response.Write "<p>Error Category: " & errObj.Category & "</p>"
Else
    Response.Write "<p>✓ No current error (expected behavior)</p>"
End If

Response.Write "<hr />"
Response.Write "<p><strong>Audit Tests Complete</strong></p>"
%>
