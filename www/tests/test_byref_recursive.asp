<%
' Test ByRef behavior in recursive function calls - validates that Application + 2D arrays work
Option Explicit

' ======== Test 1: Basic ByRef through recursive call ========
Function testByRef(ByRef s, depth)
    If depth > 0 Then
        testByRef = testByRef(s, depth - 1)
    Else
        s = "MODIFIED_BY_INNER"
        testByRef = s
    End If
End Function

Dim myStr
myStr = "ORIGINAL"
Dim result
result = testByRef(myStr, 2)

Response.Write "<h3>Test 1: ByRef through 2 levels of recursion</h3>"
Response.Write "Result: " & result & "<br>"
Response.Write "myStr after: " & myStr & "<br>"
If myStr = "MODIFIED_BY_INNER" Then
    Response.Write "<b style='color:green'>PASS: ByRef propagated through recursive calls</b><br>"
Else
    Response.Write "<b style='color:red'>FAIL: ByRef NOT propagated (got: " & myStr & ")</b><br>"
End If

Response.Write "<hr>"

' ======== Test 2: ByRef + return value assignment (mirrors treatConstants pattern) ========
Function outerFunc(ByRef s, fill)
    If InStr(1, s, "[PLACEHOLDER:") <> 0 Then
        If IsEmpty(Application("arr_cached")) Then
            ' Simulate cache initialization
            Dim arr(1, 2)
            arr(0, 0) = 1 : arr(1, 0) = "CODE1"
            arr(0, 1) = 2 : arr(1, 1) = "CODE2"
            arr(0, 2) = 3 : arr(1, 2) = "MYCODE"
            Application("arr_cached") = arr
            Application("arr_ready") = "true"
            ' Recursive call after setup (mirrors treatConstants recursive call)
            outerFunc = outerFunc(s, fill)
        End If
    End If
    ' Process replacements (mirrors insertConstants logic)
    If Not IsEmpty(Application("arr_cached")) Then
        Dim i
        For i = LBound(Application("arr_cached"), 2) To UBound(Application("arr_cached"), 2)
            Dim code
            code = Application("arr_cached")(1, i)
            If InStr(1, LCase(s), "[placeholder:" & LCase(code) & "]") <> 0 Then
                s = Replace(s, "[PLACEHOLDER:" & code & "]", "REPLACED_" & code, 1, -1, 1)
            End If
        Next
    End If
    outerFunc = s
End Function

' Reset to test fresh initialization
Application("arr_cached") = Empty
Application("arr_ready") = Empty

Dim testStr
testStr = "Hello [PLACEHOLDER:MYCODE] World"
Dim testResult
testResult = outerFunc(testStr, True)

Response.Write "<h3>Test 2: Application 2D Array + ByRef + Recursive Call (treatConstants simulation)</h3>"
Response.Write "testStr after: " & testStr & "<br>"
Response.Write "testResult: " & testResult & "<br>"
If InStr(testResult, "REPLACED_MYCODE") <> 0 Then
    Response.Write "<b style='color:green'>PASS: Placeholder correctly replaced on first call</b><br>"
Else
    Response.Write "<b style='color:red'>FAIL: Placeholder NOT replaced (got: " & testResult & ")</b><br>"
End If

Response.Write "<hr>"

' ======== Test 3: Application array - store and retrieve 2D array ========
Dim arr2d(1, 3)
arr2d(0, 0) = 10 : arr2d(1, 0) = "ALPHA"
arr2d(0, 1) = 20 : arr2d(1, 1) = "BETA"
arr2d(0, 2) = 30 : arr2d(1, 2) = "GAMMA"
arr2d(0, 3) = 40 : arr2d(1, 3) = "DELTA"

Application("test_2d_arr") = arr2d

Response.Write "<h3>Test 3: Application 2D Array - store and retrieve</h3>"
Response.Write "isEmpty: " & IsEmpty(Application("test_2d_arr")) & "<br>"
Response.Write "lbound(arr,2): " & LBound(Application("test_2d_arr"), 2) & "<br>"
Response.Write "ubound(arr,2): " & UBound(Application("test_2d_arr"), 2) & "<br>"

Dim k
For k = 0 To 3
    Response.Write "arr(0," & k & ")=" & Application("test_2d_arr")(0, k) & " | "
    Response.Write "arr(1," & k & ")=" & Application("test_2d_arr")(1, k) & "<br>"
Next

If Application("test_2d_arr")(1, 2) = "GAMMA" Then
    Response.Write "<b style='color:green'>PASS: 2D array correctly stored and retrieved from Application</b><br>"
Else
    Response.Write "<b style='color:red'>FAIL: 2D array access failed (got: " & Application("test_2d_arr")(1, 2) & ")</b><br>"
End If

Response.Write "<hr>"

' ======== Test 4: Class implicit method call (mirrors cls_customer.cacheGalleries galleries) ========
Class cls_testCache
    Public data

    Public Function getEntries()
        Set getEntries = Server.CreateObject("Scripting.Dictionary")
        getEntries.Add "A", "CODE_A"
        getEntries.Add "B", "CODE_B"
    End Function

    Public Sub buildCache()
        Dim entries
        Set entries = getEntries() ' Implicit self - method Call With ()
        Dim arr(1, 1)
        Dim runner, Key
        runner = 0
        For Each Key In entries
            arr(0, runner) = Key
            arr(1, runner) = entries(Key)
            runner = runner + 1
        Next
        data = arr
    End Sub
End Class

Dim tc
Set tc = New cls_testCache
tc.buildCache()

Response.Write "<h3>Test 4: Implicit class method call in buildCache</h3>"
If Not IsEmpty(tc.data) Then
    Response.Write "data(0,0)=" & tc.data(0, 0) & " | data(1,0)=" & tc.data(1, 0) & "<br>"
    Response.Write "data(0,1)=" & tc.data(0, 1) & " | data(1,1)=" & tc.data(1, 1) & "<br>"
    Response.Write "<b style='color:green'>PASS: Implicit class method call works</b><br>"
Else
    Response.Write "<b style='color:red'>FAIL: data is empty after buildCache</b><br>"
End If

Response.Write "<hr>"

' ======== Test 5: Set variable = function (without parentheses) ========
Class cls_dictProvider
    Public Function getDict()
        Set getDict = Server.CreateObject("Scripting.Dictionary")
        getDict.Add "key1", "val1"
    End Function

    Public Sub run()
        Dim d
        Set d = getDict ' Call without parentheses - VBScript allows this For Set
        If d.Count = 1 Then
            Response.Write "<b style='color:green'>PASS: Set d = functionName (no parens) works</b><br>"
        Else
            Response.Write "<b style='color:red'>FAIL: Set d = functionName (no parens) failed, count=" & d.Count & "</b><br>"
        End If
    End Sub
End Class

Dim dp
Set dp = New cls_dictProvider
dp.run()

Response.Write "<hr>"
Response.Write "<p>Done.</p>"
%>
