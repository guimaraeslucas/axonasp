<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include virtual="/asplite/asplite.asp"-->
<%
Response.ContentType = "text/plain"
Response.Write "=== STEP BY STEP UPLOAD TEST ===" & vbCrLf

On Error Resume Next

' Manual upload parsing following uploader.asp logic
Dim readBytes, VarArrayBinRequest, tmpBinRequest, internalChunkSize, StreamRequest

internalChunkSize = 200000

' Create stream
Set StreamRequest = Server.CreateObject("ADODB.Stream")
StreamRequest.Type = 2 'adTypeText
StreamRequest.Open

' Read binary request
readBytes = internalChunkSize
VarArrayBinRequest = Request.BinaryRead(readBytes)
Response.Write "Initial read: " & readBytes & " bytes" & vbCrLf

VarArrayBinRequest = MidB(VarArrayBinRequest, 1, LenB(VarArrayBinRequest))
Response.Write "After MidB: LenB=" & LenB(VarArrayBinRequest) & vbCrLf

' Continue reading chunks
Do Until readBytes < 1
    tmpBinRequest = Request.BinaryRead(readBytes)
    Response.Write "Chunk read: " & readBytes & " bytes" & vbCrLf
    If readBytes > 0 Then
        VarArrayBinRequest = VarArrayBinRequest & MidB(tmpBinRequest, 1, LenB(tmpBinRequest))
    End If
Loop

Response.Write "Total LenB: " & LenB(VarArrayBinRequest) & vbCrLf

' Write to stream
StreamRequest.WriteText(VarArrayBinRequest)
StreamRequest.Flush()
Response.Write "Stream Size: " & StreamRequest.Size & vbCrLf

If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Description & vbCrLf
    Err.Clear
End If

' Parse tokens
Dim tNewLine, tDoubleQuotes, tTerm, tFilename, tName, tContentDisp, tContentType
tNewLine = String2Byte(Chr(13))
tDoubleQuotes = String2Byte(Chr(34))
tTerm = String2Byte("--")
tFilename = String2Byte("filename=""")
tName = String2Byte("name=""")
tContentDisp = String2Byte("Content-Disposition")
tContentType = String2Byte("Content-Type:")

' Find first newline to get separator
Dim nCurPos
nCurPos = InstrB(1, VarArrayBinRequest, tNewLine)
Response.Write vbCrLf & "First newline at: " & nCurPos & vbCrLf

If nCurPos <= 1 Then
    Response.Write "ERROR: No newline found, exiting" & vbCrLf
    Response.End
End If

Dim vDataSep
vDataSep = MidB(VarArrayBinRequest, 1, nCurPos - 1)
Response.Write "Separator: " & vDataSep & vbCrLf

' Find last separator (terminator)
Dim nLastSepPos
nLastSepPos = InstrB(1, VarArrayBinRequest, vDataSep & tTerm)
Response.Write "Last separator at: " & nLastSepPos & vbCrLf

' Find content-disposition
nCurPos = InstrB(1, VarArrayBinRequest, tContentDisp)
Response.Write "Content-Disposition at: " & nCurPos & vbCrLf

nCurPos = InstrB(nCurPos, VarArrayBinRequest, tName)
Response.Write "name= at: " & nCurPos & vbCrLf

nCurPos = nCurPos + LenB(tName)
Dim nEndName
nEndName = InstrB(nCurPos, VarArrayBinRequest, tDoubleQuotes)
Response.Write "name ends at: " & nEndName & vbCrLf

Dim sFieldName
sFieldName = MidB(VarArrayBinRequest, nCurPos, nEndName - nCurPos)
Response.Write "Field name: " & sFieldName & vbCrLf

' Check for filename
Dim nPosFile, nPosBound
nPosFile = InstrB(nCurPos, VarArrayBinRequest, tFilename)
nPosBound = InstrB(nCurPos, VarArrayBinRequest, vDataSep)
Response.Write "filename= at: " & nPosFile & vbCrLf
Response.Write "Next boundary at: " & nPosBound & vbCrLf

If nPosFile <> 0 And nPosFile < nPosBound Then
    Response.Write "This is a FILE field!" & vbCrLf
    
    nCurPos = nPosFile + LenB(tFilename)
    Dim nEndFilename
    nEndFilename = InstrB(nCurPos, VarArrayBinRequest, tDoubleQuotes)
    
    Dim fileName
    fileName = MidB(VarArrayBinRequest, nCurPos, nEndFilename - nCurPos)
    Response.Write "File name: " & fileName & vbCrLf
    Response.Write "File type: " & aspL.getFileType(fileName) & vbCrLf
    
    ' Find content-type
    nCurPos = InstrB(nCurPos, VarArrayBinRequest, tContentType)
    Response.Write "Content-Type header at: " & nCurPos & vbCrLf
    
    nCurPos = nCurPos + LenB(tContentType)
    Dim nEndContentType
    nEndContentType = InstrB(nCurPos, VarArrayBinRequest, tNewLine)
    
    Dim contentType
    contentType = MidB(VarArrayBinRequest, nCurPos, nEndContentType - nCurPos)
    Response.Write "Content-Type: " & contentType & vbCrLf
    
    ' Skip to actual file data (skip empty line = CRLF CRLF)
    nCurPos = InstrB(nCurPos, VarArrayBinRequest, tNewLine)
    Response.Write "After content-type newline: " & nCurPos & vbCrLf
    nCurPos = nCurPos + 4 ' skip CRLF twice
    
    Dim fileStart, fileEnd, fileLength
    fileStart = nCurPos + 1
    fileEnd = InstrB(nCurPos, VarArrayBinRequest, vDataSep)
    fileLength = fileEnd - 2 - nCurPos
    
    Response.Write "File Start: " & fileStart & vbCrLf
    Response.Write "File End: " & fileEnd & vbCrLf
    Response.Write "File Length: " & fileLength & vbCrLf
    
    If fileLength > 0 Then
        Response.Write "File has content, would be added!" & vbCrLf
    Else
        Response.Write "File length <= 0, NOT added!" & vbCrLf
    End If
Else
    Response.Write "This is NOT a file field (or file before boundary)" & vbCrLf
End If

StreamRequest.Close
Set StreamRequest = Nothing

Response.Write vbCrLf & "=== DONE ===" & vbCrLf

On Error Goto 0

Function String2Byte(sString)
    Dim i
    String2Byte = ""
    For i = 1 to Len(sString)
       String2Byte = String2Byte & ChrB(AscB(Mid(sString,i,1)))
    Next
End Function
%>
