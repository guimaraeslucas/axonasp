<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.ContentType = "text/plain"
Response.Write "=== STRING2BYTE TEST ===" & vbCrLf

' Test String2Byte
Dim testStr, byteStr
testStr = "Hello"

' Using String2Byte function
byteStr = ""
Dim i
For i = 1 To Len(testStr)
    byteStr = byteStr & ChrB(AscB(Mid(testStr, i, 1)))
Next

Response.Write "testStr: " & testStr & vbCrLf
Response.Write "Len(testStr): " & Len(testStr) & vbCrLf
Response.Write "LenB(testStr): " & LenB(testStr) & vbCrLf
Response.Write vbCrLf
Response.Write "byteStr: " & byteStr & vbCrLf
Response.Write "Len(byteStr): " & Len(byteStr) & vbCrLf
Response.Write "LenB(byteStr): " & LenB(byteStr) & vbCrLf
Response.Write vbCrLf

' Compare in InstrB
Dim data
data = "Hello World"
Response.Write "data: " & data & vbCrLf
Response.Write "LenB(data): " & LenB(data) & vbCrLf

Dim pos1, pos2
pos1 = InstrB(1, data, testStr)
pos2 = InstrB(1, data, byteStr)

Response.Write "InstrB(data, testStr): " & pos1 & vbCrLf
Response.Write "InstrB(data, byteStr): " & pos2 & vbCrLf

Response.Write vbCrLf & "=== CHRB TEST ===" & vbCrLf

' Test ChrB
Dim chr13, chrStr
chr13 = ChrB(13)
Response.Write "ChrB(13):" & vbCrLf
Response.Write "  Len: " & Len(chr13) & vbCrLf
Response.Write "  LenB: " & LenB(chr13) & vbCrLf
Response.Write "  TypeName: " & TypeName(chr13) & vbCrLf

' Test MidB
Dim midResult
midResult = MidB("Hello", 1, 3)
Response.Write vbCrLf & "MidB('Hello', 1, 3):" & vbCrLf
Response.Write "  Result: " & midResult & vbCrLf
Response.Write "  Len: " & Len(midResult) & vbCrLf
Response.Write "  LenB: " & LenB(midResult) & vbCrLf
Response.Write "  TypeName: " & TypeName(midResult) & vbCrLf

Response.Write vbCrLf & "=== DONE ===" & vbCrLf
%>
