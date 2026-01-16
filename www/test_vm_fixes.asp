<%
Dim x, y
x = 10
y = 20
If x = 10 Then
    Response.Write "x is 10 (Correct)" & vbCrLf
Else
    Response.Write "x is NOT 10 (Error)" & vbCrLf
End If

If x = 20 Then
    Response.Write "x is 20 (Error)" & vbCrLf
ElseIf y = 20 Then
    Response.Write "y is 20 (Correct)" & vbCrLf
Else
    Response.Write "Neither (Error)" & vbCrLf
End If

' Test assignment vs equality
x = 5
If x = 5 Then Response.Write "Assignment worked" & vbCrLf
%>