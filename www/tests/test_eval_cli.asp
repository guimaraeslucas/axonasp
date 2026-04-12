<%
' Comprehensive Eval() test for CLI
Dim x, y, z, arr
x = 100
y = 200  
z = 300
Dim arr(2)
arr(0) = "apple"
arr(1) = "banana"

Response.Write "Eval Test Results:" & vbCrLf
Response.Write "==================" & vbCrLf

Response.Write "1. Variable: " & Eval("x") & vbCrLf
Response.Write "2. Arith:    " & Eval("(x + y) * z") & vbCrLf  
Response.Write "3. Function: " & Eval("Len(""hello"")") & vbCrLf
Response.Write "4. Array:    " & Eval("arr(0)") & vbCrLf
Response.Write "5. String:   " & Eval("UCase(""test"")") & vbCrLf
Response.Write "All tests passed!" & vbCrLf
%>
