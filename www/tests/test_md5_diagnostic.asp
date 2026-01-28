<%
Option Explicit
debug_asp_code = "TRUE"
Response.ContentType = "text/html"

Response.Write "<h3>MD5 Diagnostic Test</h3>"

' Test 1: Check if plugin loads
Response.Write "Step 1: Loading plugin...<br>"
On Error Resume Next
Dim md5Obj
Set md5Obj = aspL.plugin("md5")
If Err.Number <> 0 Then
    Response.Write "ERROR loading plugin: " & Err.Description & "<br>"
    Err.Clear
Else
    Response.Write "Plugin loaded successfully<br>"
End If
On Error Goto 0

' Test 2: Check CLng behavior
Response.Write "<br>Step 2: Testing CLng...<br>"
Dim testVal
testVal = CLng(255)
Response.Write "CLng(255) = " & testVal & " (Type: " & TypeName(testVal) & ")<br>"
testVal = CLng(2147483647)
Response.Write "CLng(2147483647) = " & testVal & "<br>"

' Test 3: Test integer division
Response.Write "<br>Step 3: Testing integer division...<br>"
Dim divResult
divResult = 10 \ 3
Response.Write "10 \ 3 = " & divResult & " (Expected: 3)<br>"

' Test 4: Test array initialization
Response.Write "<br>Step 4: Testing array...<br>"
Dim arr(5)
arr(0) = 10
arr(0) = arr(0) Or 5
Response.Write "arr(0) after Or: " & arr(0) & " (Expected: 15)<br>"

' Test 5: Try calling MD5
Response.Write "<br>Step 5: Calling MD5 function...<br>"
On Error Resume Next
Dim hashResult
hashResult = md5Obj.md5("test", 32)
If Err.Number <> 0 Then
    Response.Write "ERROR calling md5: " & Err.Description & " (Code: " & Err.Number & ")<br>"
    Err.Clear
Else
    Response.Write "MD5 returned: '" & hashResult & "' (Length: " & Len(hashResult) & ")<br>"
End If
On Error Goto 0

Response.Write "<br>Done."
%>
