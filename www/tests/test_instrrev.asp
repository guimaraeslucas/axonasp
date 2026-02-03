<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.ContentType = "text/plain"
Response.Write "=== INSTRREV TEST ===" & vbCrLf

Dim filename
filename = "Pdf.pdf"

Response.Write "Filename: " & filename & vbCrLf
Response.Write "Len(filename): " & Len(filename) & vbCrLf

Dim pos
pos = InStrRev(filename, ".")
Response.Write "InStrRev(filename, '.'): " & pos & vbCrLf

pos = InStrRev(filename, ".", -1, 1)
Response.Write "InStrRev(filename, '.', -1, 1): " & pos & vbCrLf

pos = InStrRev(filename, ".", Len(filename), 1)
Response.Write "InStrRev(filename, '.', Len, 1): " & pos & vbCrLf

If pos > 0 Then
    Dim ext
    ext = Right(filename, Len(filename) - pos)
    Response.Write "Extension: " & ext & vbCrLf
End If

Response.Write vbCrLf & "=== DONE ===" & vbCrLf
%>
