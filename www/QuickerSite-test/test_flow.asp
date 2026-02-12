<%
Response.Write "Test iID Parameter Flow" & vbCrLf
Response.Write "========================" & vbCrLf & vbCrLf

Dim testId
testId = Request("iId")

Response.Write "Step 1: Request(""iId""): [" & testId & "]" & vbCrLf
Response.Write "Step 2: IsEmpty: " & IsEmpty(testId) & vbCrLf
Response.Write "Step 3: Len: " & Len(testId) & vbCrLf
Response.Write "Step 4: Left(testId, 40): [" & Left(testId, 40) & "]" & vbCrLf

' Simulate decrypt (which likely just returns the input for invalid strings)
Function simpleDecrypt(s)
    simpleDecrypt = s
End Function

Dim decrypted
decrypted = simpleDecrypt(Left(testId, 40))
Response.Write "Step 5: After decrypt: [" & decrypted & "]" & vbCrLf

' Test if value is numeric
Function isNumeriek(val)
    On Error Resume Next
    Dim test
    test = CLng(val)
    If Err.Number = 0 Then
        isNumeriek = True
    Else
        isNumeriek = False
    End If
    On Error Goto 0
End Function

Response.Write "Step 6: isNumeriek(decrypted): " & isNumeriek(decrypted) & vbCrLf

' Test database query simulation
If isNumeriek(decrypted) Then
    Response.Write "Step 7: Would execute SQL query" & vbCrLf
Else
    Response.Write "Step 7: NO SQL query (not numeric)" & vbCrLf
End If
%>
