<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include virtual="/asplite/asplite.asp"-->
<%
Response.ContentType = "text/plain"
Response.Write "=== COMPLETE DEBUG TEST ===" & vbCrLf

On Error Resume Next

' Get binary data exactly like uploader does
Dim readBytes, VarArrayBinRequest, tmpBinRequest, internalChunkSize, StreamRequest
internalChunkSize = 200000

Set StreamRequest = Server.CreateObject("ADODB.Stream")
StreamRequest.Type = 2
StreamRequest.Open

readBytes = internalChunkSize
VarArrayBinRequest = Request.BinaryRead(readBytes)
Response.Write "Initial read: " & readBytes & " bytes, LenB=" & LenB(VarArrayBinRequest) & vbCrLf

VarArrayBinRequest = MidB(VarArrayBinRequest, 1, LenB(VarArrayBinRequest))
Response.Write "After MidB: LenB=" & LenB(VarArrayBinRequest) & ", TypeName=" & TypeName(VarArrayBinRequest) & vbCrLf

' Loop to read rest (should be done after first read)
Do Until readBytes < 1
    tmpBinRequest = Request.BinaryRead(readBytes)
    If readBytes > 0 Then
        VarArrayBinRequest = VarArrayBinRequest & MidB(tmpBinRequest, 1, LenB(tmpBinRequest))
    End If
Loop
Response.Write "Final LenB: " & LenB(VarArrayBinRequest) & vbCrLf

StreamRequest.WriteText(VarArrayBinRequest)
StreamRequest.Flush()
Response.Write "Stream Size: " & StreamRequest.Size & vbCrLf

' Create tokens
Dim tNewLine, tDoubleQuotes, tTerm, tFilename, tName, tContentDisp
tNewLine = ChrB(13)
tDoubleQuotes = ChrB(34)
tTerm = ChrB(45) & ChrB(45)  ' "--"
tFilename = ""
Dim fi
For fi = 1 To Len("filename=""")
    tFilename = tFilename & ChrB(AscB(Mid("filename=""", fi, 1)))
Next
tName = ""
For fi = 1 To Len("name=""")
    tName = tName & ChrB(AscB(Mid("name=""", fi, 1)))
Next
tContentDisp = ""
For fi = 1 To Len("Content-Disposition")
    tContentDisp = tContentDisp & ChrB(AscB(Mid("Content-Disposition", fi, 1)))
Next

Response.Write vbCrLf & "Token info:" & vbCrLf
Response.Write "  tNewLine LenB: " & LenB(tNewLine) & ", TypeName: " & TypeName(tNewLine) & vbCrLf
Response.Write "  tFilename LenB: " & LenB(tFilename) & ", TypeName: " & TypeName(tFilename) & vbCrLf
Response.Write "  tContentDisp LenB: " & LenB(tContentDisp) & ", TypeName: " & TypeName(tContentDisp) & vbCrLf

' Test searching
Response.Write vbCrLf & "Search tests:" & vbCrLf
Dim pos
pos = InstrB(1, VarArrayBinRequest, tNewLine)
Response.Write "  Newline position: " & pos & vbCrLf

If pos > 1 Then
    Dim vDataSep
    vDataSep = MidB(VarArrayBinRequest, 1, pos - 1)
    Response.Write "  Separator LenB: " & LenB(vDataSep) & vbCrLf
    
    pos = InstrB(1, VarArrayBinRequest, tContentDisp)
    Response.Write "  Content-Disposition position: " & pos & vbCrLf
    
    pos = InstrB(1, VarArrayBinRequest, tFilename)
    Response.Write "  filename= position: " & pos & vbCrLf
    
    ' Try to find with string
    Dim tFilenameStr
    tFilenameStr = "filename="""
    pos = InstrB(1, VarArrayBinRequest, tFilenameStr)
    Response.Write "  filename= (string) position: " & pos & vbCrLf
End If

If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Description & vbCrLf
    Err.Clear
End If

StreamRequest.Close
Set StreamRequest = Nothing

Response.Write vbCrLf & "=== DONE ===" & vbCrLf
On Error Goto 0
%>
