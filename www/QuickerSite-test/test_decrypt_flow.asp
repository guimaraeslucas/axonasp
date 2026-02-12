<!-- #include file="asp/includes/constants.inc" -->
<!-- #include file="asp/includes/functions.asp" -->
<!-- #include file="asp/includes/encryption.asp" -->
<%
Response.Write "Test Full Decrypt Flow with iID" & vbCrLf
Response.Write "==================================" & vbCrLf & vbCrLf

Dim testId
testId = Request("iId")

Response.Write "Step 1: Request(""iId""): [" & testId & "]" & vbCrLf
Response.Write "Step 2: Left(testId, 40): [" & Left(testId, 40) & "]" & vbCrLf

Dim decrypted
On Error Resume Next
decrypted = decrypt(Left(testId, 40))
If Err.Number <> 0 Then
    Response.Write "ERROR in decrypt: " & Err.Description & " (Error " & Err.Number & ")" & vbCrLf
    Err.Clear
End If
On Error Goto 0

Response.Write "Step 3: After decrypt: [" & decrypted & "]" & vbCrLf
Response.Write "Step 4: isNumeriek(decrypted): " & isNumeriek(decrypted) & vbCrLf

Response.Write vbCrLf & "Test complete - no crashes!" & vbCrLf
%>
