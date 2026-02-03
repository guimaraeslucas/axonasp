<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.ContentType = "text/plain"
Response.Write "=== BINARY FUNCTIONS TEST ===" & vbCrLf

On Error Resume Next

' Test InstrB
Dim testStr
testStr = "Hello World"
Response.Write "testStr: " & testStr & vbCrLf
Response.Write "LenB(testStr): " & LenB(testStr) & vbCrLf
Response.Write "InstrB(1, testStr, 'o'): " & InstrB(1, testStr, "o") & vbCrLf
Response.Write "InstrB(1, testStr, 'World'): " & InstrB(1, testStr, "World") & vbCrLf
Response.Write "MidB(testStr, 7, 5): " & MidB(testStr, 7, 5) & vbCrLf
Response.Write vbCrLf

' Now test with binary data from BinaryRead
Dim readBytes, binData
readBytes = 500
binData = Request.BinaryRead(readBytes)

Response.Write "Bytes read: " & readBytes & vbCrLf
Response.Write "LenB(binData): " & LenB(binData) & vbCrLf

' Search for boundary string
Dim boundary
boundary = "Content-Disposition"
Response.Write "Looking for: " & boundary & vbCrLf
Response.Write "LenB(boundary): " & LenB(boundary) & vbCrLf

Dim pos
pos = InstrB(1, binData, boundary)
Response.Write "InstrB found at position: " & pos & vbCrLf

' Try with String2Byte style
Dim i, binaryBoundary
binaryBoundary = ""
For i = 1 To Len(boundary)
    binaryBoundary = binaryBoundary & ChrB(AscB(Mid(boundary, i, 1)))
Next
Response.Write "LenB(binaryBoundary): " & LenB(binaryBoundary) & vbCrLf

pos = InstrB(1, binData, binaryBoundary)
Response.Write "InstrB with binary boundary at position: " & pos & vbCrLf

' Try to extract from binData
If pos > 0 Then
    Dim extracted
    extracted = MidB(binData, pos, 19)
    Response.Write "Extracted: " & extracted & vbCrLf
End If

Response.Write "=== DONE ===" & vbCrLf

On Error Goto 0
%>
