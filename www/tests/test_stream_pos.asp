<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.ContentType = "text/plain"
Response.Write "=== STREAM POSITION DEBUG ===" & vbCrLf

On Error Resume Next

Dim StreamRequest, VarArrayBinRequest, readBytes

' Read binary data
readBytes = 200000
VarArrayBinRequest = Request.BinaryRead(readBytes)
Response.Write "Read " & readBytes & " bytes" & vbCrLf
Response.Write "LenB(VarArrayBinRequest): " & LenB(VarArrayBinRequest) & vbCrLf

' Write to stream
Set StreamRequest = Server.CreateObject("ADODB.Stream")
StreamRequest.Type = 2
StreamRequest.Open
StreamRequest.WriteText(VarArrayBinRequest)
StreamRequest.Flush()

Response.Write "Stream Size: " & StreamRequest.Size & vbCrLf

' Now search for filename in the original data
Dim tFilename, i
tFilename = ""
For i = 1 to Len("filename=""")
   tFilename = tFilename & ChrB(AscB(Mid("filename=""",i,1)))
Next
Response.Write "tFilename LenB: " & LenB(tFilename) & vbCrLf

Dim nPosFile
nPosFile = InstrB(1, VarArrayBinRequest, tFilename)
Response.Write "filename= at position (InstrB): " & nPosFile & vbCrLf

' After tFilename, we should be at the start of the actual filename
Dim startPos
startPos = nPosFile + LenB(tFilename)
Response.Write "Filename starts at (after filename=): " & startPos & vbCrLf

' Find the closing quote
Dim tQuote
tQuote = ChrB(34)
Dim endPos
endPos = InstrB(startPos, VarArrayBinRequest, tQuote)
Response.Write "Closing quote at: " & endPos & vbCrLf

' Calculate length
Dim fnLen
fnLen = endPos - startPos
Response.Write "Filename length: " & fnLen & vbCrLf

' Try to read from original binary data directly
Dim fnFromBinary
fnFromBinary = MidB(VarArrayBinRequest, startPos, fnLen)
Response.Write "Filename from MidB (binary): [" & fnFromBinary & "]" & vbCrLf

' Now try to read from stream
' The uploader uses: Position = start + 1
Dim streamPos
streamPos = startPos + 1  ' As the uploader does: start + 1
Response.Write vbCrLf & "Setting stream position to: " & streamPos & " (using start+1 like uploader)" & vbCrLf
StreamRequest.Position = streamPos

' Create dest stream
Dim dstStream
Set dstStream = Server.CreateObject("ADODB.Stream")
dstStream.Charset = "utf-8"
dstStream.Mode = 3
dstStream.Type = 1
dstStream.Open

' Copy
StreamRequest.CopyTo dstStream, fnLen
dstStream.Flush

Response.Write "Dest size: " & dstStream.Size & vbCrLf

' Read
dstStream.Position = 0
dstStream.Type = 2
Dim fnFromStream
fnFromStream = dstStream.ReadText
Response.Write "Filename from stream: [" & fnFromStream & "]" & vbCrLf

dstStream.Close
StreamRequest.Close

If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf
    Err.Clear
End If

Response.Write vbCrLf & "=== DONE ===" & vbCrLf
%>
