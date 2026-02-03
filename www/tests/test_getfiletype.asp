<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include virtual="/asplite/asplite.asp"-->
<%
Response.ContentType = "text/plain"
Response.Write "=== getFileType TEST ===" & vbCrLf

Dim fn
fn = "Pdf.pdf"
Response.Write "Filename: " & fn & vbCrLf

' Test Len
Response.Write "Len: " & Len(fn) & vbCrLf

' Test InStrRev with -1 start
Dim instrrev_pos
instrrev_pos = InStrRev(fn, ".", -1, 1)
Response.Write "InStrRev(fn, '.', -1, 1): " & instrrev_pos & vbCrLf

' Test InStrRev with 0 start 
instrrev_pos = InStrRev(fn, ".")
Response.Write "InStrRev(fn, '.'): " & instrrev_pos & vbCrLf

' Test calculation
Dim calc
calc = Len(fn) - InStrRev(fn, ".", -1, 1)
Response.Write "Len - InStrRev: " & calc & vbCrLf

' Test right
Dim ft
ft = right(fn, Len(fn) - InStrRev(fn, ".", -1, 1))
Response.Write "Right result: " & ft & vbCrLf

' Test aspL function
Response.Write "aspL.getFileType: " & aspL.getFileType(fn) & vbCrLf

' Test LCase
Response.Write "LCase(FileType): " & LCase(ft) & vbCrLf

Response.Write vbCrLf & "=== DONE ===" & vbCrLf
%>
