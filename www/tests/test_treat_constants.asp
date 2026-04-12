<%@ Language="VBScript" %>
<%
' Diagnostic test that simulates QuickerSite's treatConstants pipeline
' This tests the exact pattern that is used to replace [MENU], [LOGO], etc.

Response.Write "<html><body>"
Response.Write "<h1>treatConstants Pipeline Diagnostic</h1>"

Dim passed, failed
passed = 0
failed = 0

Sub AssertEqual(actual, expected, desc)
    If CStr(actual) = CStr(expected) Then
        Response.Write "<div style='color:green'>PASS: " & desc & " (got: " & Server.HTMLEncode(CStr(actual)) & ")</div>" & vbCrLf
        passed = passed + 1
    Else
        Response.Write "<div style='color:red'>FAIL: " & desc & " - Expected: " & Server.HTMLEncode(CStr(expected)) & ", Got: " & Server.HTMLEncode(CStr(actual)) & "</div>" & vbCrLf
        failed = failed + 1
    End If
End Sub

' ============================================================
' Simulate QuickerSite constant caching
' ============================================================
Response.Write "<h2>Phase 1: Cache Constants</h2>"

Dim arrconstants
ReDim arrconstants(2, 3)

' Constant 0: MENU (text/html type)
arrconstants(0, 0) = "MENU"
arrconstants(1, 0) = "<nav>Home | About | Contact</nav>"
arrconstants(2, 0) = ""

' Constant 1: LOGO (text/html type)
arrconstants(0, 1) = "LOGO"
arrconstants(1, 1) = "<img src='logo.png' />"
arrconstants(2, 1) = ""

' Constant 2: ADDRESS (text/html type)
arrconstants(0, 2) = "ADDRESS"
arrconstants(1, 2) = "123 Main St"
arrconstants(2, 2) = ""

' Constant 3: SITEINFO (VBScript type - with identifier)
arrconstants(0, 3) = "SITEINFO"
arrconstants(1, 3) = "Response.Write ""test_value""|||QS_VBS|||param1"
arrconstants(2, 3) = ""

' Store in Application
Application.Lock
Application("test_arrconstants") = arrconstants
Application.Unlock

' Verify storage
AssertEqual Application("test_arrconstants")(0, 0), "MENU", "App stored: MENU name"
AssertEqual Application("test_arrconstants")(1, 0), "<nav>Home | About | Contact</nav>", "App stored: MENU value"
AssertEqual Application("test_arrconstants")(0, 1), "LOGO", "App stored: LOGO name"
AssertEqual Application("test_arrconstants")(0, 2), "ADDRESS", "App stored: ADDRESS name"

' Verify LBound/UBound for dimension 2
AssertEqual LBound(Application("test_arrconstants"), 1), 0, "LBound dim 1"
AssertEqual UBound(Application("test_arrconstants"), 1), 2, "UBound dim 1"
AssertEqual LBound(Application("test_arrconstants"), 2), 0, "LBound dim 2"
AssertEqual UBound(Application("test_arrconstants"), 2), 3, "UBound dim 2"

' ============================================================
' Simulate treatConstants replacement (text/html type)
' ============================================================
Response.Write "<h2>Phase 2: Constant Replacement (Text/HTML)</h2>"

Dim origValue, cname, re, iconstantKey
origValue = "Header: [LOGO] Content: [MENU] Footer: [ADDRESS]"

Response.Write "<div>Input: " & Server.HTMLEncode(origValue) & "</div>"

If InStr(1, origValue, "[", vbTextCompare) <> 0 Then
    Set re = New RegExp
    re.Global = True
    re.IgnoreCase = True
    
    For iconstantKey = LBound(Application("test_arrconstants"), 2) To UBound(Application("test_arrconstants"), 2)
        cname = Application("test_arrconstants")(0, iconstantKey)
        
        ' Simple text/html replacement (no VBScript identifier)
        If InStr(1, Application("test_arrconstants")(1, iconstantKey), "|||QS_VBS|||", vbTextCompare) = 0 Then
            re.Pattern = "\[+" & cname & "+\]"
            If re.Test(origValue) Then
                origValue = Replace(origValue, "[" & cname & "]", Application("test_arrconstants")(1, iconstantKey), 1, -1, 1)
            End If
        End If
    Next
    
    Set re = Nothing
End If

Response.Write "<div>Output: " & Server.HTMLEncode(origValue) & "</div>"

AssertEqual InStr(origValue, "[MENU]"), 0, "No [MENU] placeholder remaining"
AssertEqual InStr(origValue, "[LOGO]"), 0, "No [LOGO] placeholder remaining"
AssertEqual InStr(origValue, "[ADDRESS]"), 0, "No [ADDRESS] placeholder remaining"
AssertEqual InStr(origValue, "<nav>"), 1, "Contains nav element"

' ============================================================
' Test with RegExp Execute + For Each (VBScript-type constants)
' ============================================================
Response.Write "<h2>Phase 3: RegExp Execute + For Each</h2>"

Dim testStr, matches, mv
testStr = "Hello [SITEINFO(""name"")] World"

Set re = New RegExp
re.Global = True
re.IgnoreCase = True
re.Pattern = "\[+(SITEINFO)+[\(]+[\S| ]+[\)]+[\]]|\[+(SITEINFO)+[\]]"

If re.Test(testStr) Then
    Response.Write "<div>Pattern matched in: " & Server.HTMLEncode(testStr) & "</div>"
    
    Set matches = re.Execute(Replace(testStr, "]", "]" & vbCrLf, 1, -1, 1))
    AssertEqual matches.Count > 0, True, "Execute found matches"
    
    Dim matchValues
    matchValues = ""
    For Each mv In matches
        matchValues = matchValues & mv.Value & "|"
        Response.Write "<div>Match: " & Server.HTMLEncode(mv.Value) & "</div>"
    Next
    
    AssertEqual Len(matchValues) > 0, True, "For Each iterated over matches"
Else
    Response.Write "<div style='color:red'>Pattern did NOT match!</div>"
    failed = failed + 1
End If

Set re = Nothing

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
