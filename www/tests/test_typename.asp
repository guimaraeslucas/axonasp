<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include virtual="/asplite/asplite.asp"-->
<%
Response.ContentType = "text/plain"
Response.Write "=== TYPENAME DEBUG ===" & vbCrLf

On Error Resume Next

Dim readBytes, VarArrayBinRequest
readBytes = 200000
VarArrayBinRequest = Request.BinaryRead(readBytes)

Response.Write "readBytes: " & readBytes & vbCrLf
Response.Write "LenB: " & LenB(VarArrayBinRequest) & vbCrLf
Response.Write "TypeName: " & TypeName(VarArrayBinRequest) & vbCrLf

' Test various type checks
Response.Write vbCrLf & "Type checks:" & vbCrLf
Response.Write "  IsEmpty: " & IsEmpty(VarArrayBinRequest) & vbCrLf
Response.Write "  IsNull: " & IsNull(VarArrayBinRequest) & vbCrLf
Response.Write "  VarType: " & VarType(VarArrayBinRequest) & vbCrLf

' Get substring
Dim sub1
sub1 = MidB(VarArrayBinRequest, 1, 10)
Response.Write vbCrLf & "MidB result:" & vbCrLf
Response.Write "  LenB: " & LenB(sub1) & vbCrLf
Response.Write "  TypeName: " & TypeName(sub1) & vbCrLf
Response.Write "  VarType: " & VarType(sub1) & vbCrLf

' Get using ChrB
Dim chrb1
chrb1 = ChrB(65)
Response.Write vbCrLf & "ChrB result:" & vbCrLf
Response.Write "  LenB: " & LenB(chrb1) & vbCrLf
Response.Write "  TypeName: " & TypeName(chrb1) & vbCrLf 
Response.Write "  VarType: " & VarType(chrb1) & vbCrLf

' Concatenation
Dim concat1
concat1 = ChrB(65) & ChrB(66)
Response.Write vbCrLf & "Concat result:" & vbCrLf
Response.Write "  LenB: " & LenB(concat1) & vbCrLf
Response.Write "  TypeName: " & TypeName(concat1) & vbCrLf
Response.Write "  VarType: " & VarType(concat1) & vbCrLf

' Test if the "Unknown" type works with InstrB
Dim pos
pos = InstrB(1, VarArrayBinRequest, concat1)
Response.Write vbCrLf & "InstrB test (search AB in binary data):" & vbCrLf
Response.Write "  Position: " & pos & vbCrLf

' Test Select Case with numbers
Dim testNum
testNum = 1
Response.Write vbCrLf & "Select Case test:" & vbCrLf
Select Case testNum
    Case 0
        Response.Write "  Matched 0" & vbCrLf
    Case 1
        Response.Write "  Matched 1" & vbCrLf
    Case 2
        Response.Write "  Matched 2" & vbCrLf
    Case Else
        Response.Write "  Matched Else" & vbCrLf
End Select

' Test binary comparison equality
Dim cmp1, cmp2
cmp1 = ChrB(65) & ChrB(66)
cmp2 = ChrB(65) & ChrB(66)
Response.Write vbCrLf & "Binary comparison:" & vbCrLf
Response.Write "  cmp1 = cmp2: " & (cmp1 = cmp2) & vbCrLf

If Err.Number <> 0 Then
    Response.Write vbCrLf & "ERROR: " & Err.Description & vbCrLf
    Err.Clear
End If

Response.Write vbCrLf & "=== DONE ===" & vbCrLf
%>
