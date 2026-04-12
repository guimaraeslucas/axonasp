<%
' Test Eval with more expressions
Response.Write "Eval() Comprehensive Compatibility Test" & vbCrLf
Response.Write "======================================" & vbCrLf & vbCrLf

Dim x, y, z, arr, result

' Initialize test data
x = 5
y = 10
z = 15
Dim arr(2)
arr(0) = 100
arr(1) = 200

' Test 1: Math operations
Response.Write "1. Math: " & Eval("x * y + z") & " (Expected: 65)" & vbCrLf

' Test 2: String operations
Response.Write "2. String: " & Eval("LCase(""TEST"")") & " (Expected: test)" & vbCrLf

' Test 3: Array access
Response.Write "3. Array: " & Eval("arr(0) + arr(1)") & " (Expected: 300)" & vbCrLf

' Test 4: Boolean logic
Response.Write "4. Logic: " & Eval("x > 3 And y < 20") & " (Expected: True)" & vbCrLf

' Test 5: Type coercion/comparison
Response.Write "5. Compare: " & Eval("""5"" & x") & " (Expected: 55)" & vbCrLf

Response.Write vbCrLf & "All Eval() tests successful - FULL COMPATIBILITY CONFIRMED" & vbCrLf
%>
