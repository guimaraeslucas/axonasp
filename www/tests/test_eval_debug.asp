<%
Response.Write("=== EVAL() DEBUGGING ===" & vbCrLf & vbCrLf)

' Set up test variables
Dim x, y, z
x = 100
y = 200
z = 300

Response.Write("Variables set: x=" & x & ", y=" & y & ", z=" & z & vbCrLf & vbCrLf)

' Test 1: Direct assignment (control test)
Response.Write("Test 1 - Direct Assignment (Control):" & vbCrLf)
Dim temp1
temp1 = x
Response.Write("  temp1 = x  -> temp1 = " & temp1 & vbCrLf)
Response.Write(vbCrLf)

' Test 2: Assignment with arithmetic (control test)
Response.Write("Test 2 - Arithmetic Expression (Control):" & vbCrLf)
Dim temp2
temp2 = x + 5
Response.Write("  temp2 = x + 5  -> temp2 = " & temp2 & vbCrLf)
Response.Write(vbCrLf)

' Test 3: Eval() with simple variable
Response.Write("Test 3 - Eval() with Variable:" & vbCrLf)
Dim result3
result3 = Eval("x")
Dim msg3
If result3 = "" Then
    msg3 = "EMPTY"
Else
    msg3 = result3
End If
Response.Write("  result3 = Eval(""x"")  -> result3 = " & result3 & vbCrLf)
Response.Write("  Expected: 100 | Got: " & msg3 & vbCrLf)
Response.Write(vbCrLf)

' Test 4: Eval() with arithmetic
Response.Write("Test 4 - Eval() with Arithmetic:" & vbCrLf)
Dim result4
result4 = Eval("x + 5")
Dim msg4
If result4 = "" Then
    msg4 = "EMPTY"
Else
    msg4 = result4
End If
Response.Write("  result4 = Eval(""x + 5"")  -> result4 = " & result4 & vbCrLf)
Response.Write("  Expected: 105 | Got: " & msg4 & vbCrLf)
Response.Write(vbCrLf)

' Test 5: Eval() with complex expression
Response.Write("Test 5 - Eval() with Complex Expression:" & vbCrLf)
Dim result5
result5 = Eval("(x + y) * z")
Dim msg5
If result5 = "" Then
    msg5 = "EMPTY"
Else
    msg5 = result5
End If
Response.Write("  result5 = Eval(""(x + y) * z"")  -> result5 = " & result5 & vbCrLf)
Response.Write("  Expected: 90000 | Got: " & msg5 & vbCrLf)
Response.Write(vbCrLf)

' Test 6: Direct evaluation of complex expression
Response.Write("Test 6 - Direct Evaluation (Control):" & vbCrLf)
Dim result6
result6 = (x + y) * z
Response.Write("  result6 = (x + y) * z  -> result6 = " & result6 & vbCrLf)
Response.Write(vbCrLf)

' Test 7: Eval() function call
Response.Write("Test 7 - Eval() with Function Call:" & vbCrLf)
Dim result7
result7 = Eval("Len(""hello"")")
Dim msg7
If result7 = "" Then
    msg7 = "EMPTY"
Else
    msg7 = result7
End If
Response.Write("  result7 = Eval(""Len(""hello"")"")  -> result7 = " & result7 & vbCrLf)
Response.Write("  Expected: 5 | Got: " & msg7 & vbCrLf)
Response.Write(vbCrLf)

Response.Write("=== END DEBUG ===" & vbCrLf)
%>
