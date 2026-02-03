<%@ Language="VBScript" %>
<%
' Test InstrB with binary data to trace multipart parsing
Response.ContentType = "text/plain"

' Read multipart data
Dim nTotalBytes
nTotalBytes = Request.TotalBytes
Response.Write "TotalBytes: " & nTotalBytes & vbCrLf

If nTotalBytes > 0 Then
    Dim VarArrayBinRequest
    VarArrayBinRequest = Request.BinaryRead(nTotalBytes)
    
    Response.Write "Data type: " & TypeName(VarArrayBinRequest) & vbCrLf
    Response.Write "Data length (LenB): " & LenB(VarArrayBinRequest) & vbCrLf
    
    ' Convert to string to show content
    Dim strData
    strData = ""
    Dim i
    For i = 1 To LenB(VarArrayBinRequest)
        Dim b
        b = AscB(MidB(VarArrayBinRequest, i, 1))
        If b >= 32 And b < 127 Then
            strData = strData & Chr(b)
        ElseIf b = 13 Then
            strData = strData & "[CR]"
        ElseIf b = 10 Then
            strData = strData & "[LF]"
        Else
            strData = strData & "[" & b & "]"
        End If
    Next
    Response.Write "Content preview:" & vbCrLf & strData & vbCrLf & vbCrLf
    
    ' Test InstrB with the boundary
    Dim boundary
    boundary = ChrB(45) & ChrB(45) & ChrB(45) & ChrB(45) & ChrB(45) & ChrB(45) & ChrB(84) & ChrB(101) & ChrB(115) & ChrB(116) & ChrB(66) & ChrB(111) & ChrB(117) & ChrB(110) & ChrB(100) & ChrB(97) & ChrB(114) & ChrB(121) & ChrB(49) & ChrB(50) & ChrB(51) & ChrB(52) & ChrB(53)
    Response.Write "Looking for boundary (LenB): " & LenB(boundary) & vbCrLf
    
    Response.Write vbCrLf & "Testing InstrB:" & vbCrLf
    
    ' Find first occurrence
    Dim pos1
    pos1 = InstrB(1, VarArrayBinRequest, boundary)
    Response.Write "First boundary at position: " & pos1 & vbCrLf
    
    ' Find second occurrence
    If pos1 > 0 Then
        Dim pos2
        pos2 = InstrB(pos1 + 1, VarArrayBinRequest, boundary)
        Response.Write "Second boundary at position: " & pos2 & vbCrLf
    End If
    
    ' Look for CRLF
    Dim crlf
    crlf = ChrB(13) & ChrB(10)
    Dim posCRLF
    posCRLF = InstrB(1, VarArrayBinRequest, crlf)
    Response.Write "First CRLF at position: " & posCRLF & vbCrLf
    
    ' Look for double CRLF (end of headers)
    Dim doubleCRLF
    doubleCRLF = ChrB(13) & ChrB(10) & ChrB(13) & ChrB(10)
    Dim posDoubleCRLF
    posDoubleCRLF = InstrB(1, VarArrayBinRequest, doubleCRLF)
    Response.Write "Double CRLF (header end) at position: " & posDoubleCRLF & vbCrLf
    
Else
    Response.Write "No data received. Send a multipart/form-data POST request."
End If
%>
