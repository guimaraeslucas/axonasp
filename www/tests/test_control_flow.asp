<%
Response.Write("Test 1: If statement" & vbCrLf)
If 1 > 0 Then
    Response.Write("If works!" & vbCrLf)
End If
Response.Write("Test 2: If-Else" & vbCrLf)
x = 5
If x > 10 Then
    Response.Write("x > 10" & vbCrLf)
Else
    Response.Write("x <= 10" & vbCrLf)
End If
Response.Write("Test 3: For loop" & vbCrLf)
For i = 1 To 3
    Response.Write("i=" & i & vbCrLf)
Next
Response.Write("Test 4: Array" & vbCrLf)
Dim arr(2)
arr(0) = "a"
arr(1) = "b"
arr(2) = "c"
Response.Write("arr(0)=" & arr(0) & vbCrLf)
%>
