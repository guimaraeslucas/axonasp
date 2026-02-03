<%@ Language="VBScript" %>
<%
Response.ContentType = "text/plain"
Response.Write "=== TEST INSTRREV ===" & vbCrLf

Dim filename
filename = "test.txt"

Response.Write "Filename: " & filename & vbCrLf
Response.Write "Len: " & Len(filename) & vbCrLf

Dim dotPos
dotPos = InStrRev(filename, ".", -1, 1)
Response.Write "InStrRev result: " & dotPos & vbCrLf

If dotPos > 0 Then
    Dim ext
    ext = Right(filename, Len(filename) - dotPos)
    Response.Write "Extension: " & ext & vbCrLf
Else
    Response.Write "No dot found" & vbCrLf
End If

Response.Write vbCrLf & "=== DONE ===" & vbCrLf
%>
