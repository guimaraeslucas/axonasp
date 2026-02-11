<%@ Language="VBScript" %>
<%
' Test multi-dimensional array assignment and RegExp For Each iteration
' These are critical for QuickerSite's constant replacement pipeline

Dim passed, failed, testName
passed = 0
failed = 0

Sub AssertEqual(actual, expected, desc)
    If CStr(actual) = CStr(expected) Then
        Response.Write "<div style='color:green'>PASS: " & desc & " (got: " & actual & ")</div>" & vbCrLf
        passed = passed + 1
    Else
        Response.Write "<div style='color:red'>FAIL: " & desc & " - Expected: " & expected & ", Got: " & actual & "</div>" & vbCrLf
        failed = failed + 1
    End If
End Sub

Sub AssertTrue(val, desc)
    If val Then
        Response.Write "<div style='color:green'>PASS: " & desc & "</div>" & vbCrLf
        passed = passed + 1
    Else
        Response.Write "<div style='color:red'>FAIL: " & desc & "</div>" & vbCrLf
        failed = failed + 1
    End If
End Sub

Response.Write "<html><body>"
Response.Write "<h1>Multi-Dimensional Array and RegExp Tests</h1>"

' ============================================================
' TEST 1: Multi-dimensional array assignment and read
' ============================================================
Response.Write "<h2>Test 1: Multi-Dimensional Array Assignment</h2>"

Dim arr
ReDim arr(2, 3)

arr(0, 0) = "r0c0"
arr(0, 1) = "r0c1"
arr(0, 2) = "r0c2"
arr(0, 3) = "r0c3"
arr(1, 0) = "r1c0"
arr(1, 1) = "r1c1"
arr(2, 0) = "r2c0"
arr(2, 3) = "r2c3"

AssertEqual arr(0, 0), "r0c0", "arr(0,0)"
AssertEqual arr(0, 1), "r0c1", "arr(0,1)"
AssertEqual arr(0, 2), "r0c2", "arr(0,2)"
AssertEqual arr(0, 3), "r0c3", "arr(0,3)"
AssertEqual arr(1, 0), "r1c0", "arr(1,0)"
AssertEqual arr(1, 1), "r1c1", "arr(1,1)"
AssertEqual arr(2, 0), "r2c0", "arr(2,0)"
AssertEqual arr(2, 3), "r2c3", "arr(2,3)"

' ============================================================
' TEST 2: Simulating QuickerSite's arrconstants cache
' ============================================================
Response.Write "<h2>Test 2: QuickerSite arrconstants Simulation</h2>"

Dim arrconstants
ReDim arrconstants(2, 4)

' Simulate caching constants like QuickerSite does
arrconstants(0, 0) = "SITEINFO"
arrconstants(1, 0) = "Response.Write GetSiteInfo(param)"
arrconstants(2, 0) = ""

arrconstants(0, 1) = "MENU"
arrconstants(1, 1) = "Response.Write BuildMenu()"
arrconstants(2, 1) = ""

arrconstants(0, 2) = "SITENAME"
arrconstants(1, 2) = "My Website"
arrconstants(2, 2) = "direct"

AssertEqual arrconstants(0, 0), "SITEINFO", "arrconstants(0,0) = SITEINFO"
AssertEqual arrconstants(1, 0), "Response.Write GetSiteInfo(param)", "arrconstants(1,0) = code"
AssertEqual arrconstants(0, 1), "MENU", "arrconstants(0,1) = MENU"
AssertEqual arrconstants(1, 1), "Response.Write BuildMenu()", "arrconstants(1,1) = code"
AssertEqual arrconstants(0, 2), "SITENAME", "arrconstants(0,2) = SITENAME"
AssertEqual arrconstants(2, 2), "direct", "arrconstants(2,2) = direct"

' ============================================================
' TEST 3: Store/retrieve 2D array in Application
' ============================================================
Response.Write "<h2>Test 3: 2D Array in Application State</h2>"

Application.Lock
Application("test_arr") = arrconstants
Application.Unlock

Dim retrieved
Set retrieved = Nothing
Dim retVal
retVal = Application("test_arr")

' Access the array from Application state
AssertEqual retVal(0, 0), "SITEINFO", "Application arr (0,0)"
AssertEqual retVal(0, 1), "MENU", "Application arr (0,1)"
AssertEqual retVal(1, 2), "My Website", "Application arr (1,2)"

' ============================================================
' TEST 4: RegExp For Each iteration
' ============================================================
Response.Write "<h2>Test 4: RegExp For Each Iteration</h2>"

Dim re, matches, m, matchCount
Set re = New RegExp
re.Pattern = "\d+"
re.Global = True

Set matches = re.Execute("abc 123 def 456 ghi 789")

AssertEqual matches.Count, 3, "RegExp match count"

matchCount = 0
For Each m In matches
    matchCount = matchCount + 1
    If matchCount = 1 Then AssertEqual m.Value, "123", "Match 1 value"
    If matchCount = 2 Then AssertEqual m.Value, "456", "Match 2 value"
    If matchCount = 3 Then AssertEqual m.Value, "789", "Match 3 value"
Next

AssertEqual matchCount, 3, "For Each iterated 3 times"

' ============================================================
' TEST 5: InStr with compare modes
' ============================================================
Response.Write "<h2>Test 5: InStr Compare Modes</h2>"

' Default (binary) - case sensitive
AssertEqual InStr("Hello World", "hello"), 0, "InStr binary: 'hello' not in 'Hello World'"
AssertEqual InStr("Hello World", "Hello"), 1, "InStr binary: 'Hello' in 'Hello World'"

' Text compare - case insensitive
AssertEqual InStr(1, "Hello World", "hello", 1), 1, "InStr text: 'hello' in 'Hello World'"

' Binary compare explicit
AssertEqual InStr(1, "Hello World", "hello", 0), 0, "InStr binary explicit: 'hello' not in 'Hello World'"

' ============================================================
' TEST 6: Replace with compare modes
' ============================================================
Response.Write "<h2>Test 6: Replace Compare Modes</h2>"

' Default (binary) - case sensitive
AssertEqual Replace("Hello World", "hello", "HI"), "Hello World", "Replace binary: no match for 'hello'"
AssertEqual Replace("Hello World", "Hello", "HI"), "HI World", "Replace binary: 'Hello' -> 'HI'"

' Text compare
AssertEqual Replace("Hello World", "hello", "HI", 1, -1, 1), "HI World", "Replace text: 'hello' -> 'HI'"

' Replace bracket tags (QuickerSite pattern)
Dim template
template = "Welcome to [SITEINFO(""name"")]!"
AssertEqual Replace(template, "[SITEINFO(""name"")]", "My Site"), "Welcome to My Site!", "Replace bracket tag"

' ============================================================
' TEST 7: VBScript compare constants
' ============================================================
Response.Write "<h2>Test 7: VBScript Compare Constants</h2>"

AssertEqual vbBinaryCompare, 0, "vbBinaryCompare = 0"
AssertEqual vbTextCompare, 1, "vbTextCompare = 1"

' Use constants in InStr
AssertEqual InStr(1, "Hello", "hello", vbTextCompare), 1, "InStr with vbTextCompare"
AssertEqual InStr(1, "Hello", "hello", vbBinaryCompare), 0, "InStr with vbBinaryCompare"

' ============================================================
' SUMMARY
' ============================================================
Response.Write "<hr>"
Response.Write "<h2>Results: " & passed & " passed, " & failed & " failed</h2>"
If failed = 0 Then
    Response.Write "<div style='color:green;font-size:24px'>ALL TESTS PASSED</div>"
Else
    Response.Write "<div style='color:red;font-size:24px'>SOME TESTS FAILED</div>"
End If

Response.Write "</body></html>"
%>
