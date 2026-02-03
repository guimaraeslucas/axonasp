<%@ Language="VBScript" %>
<%
Response.ContentType = "text/html"

' Check request
If Request.TotalBytes < 1 Then
    Response.Write "<p>No data. POST multipart data to this endpoint.</p>"
    Response.End
End If

Response.Write "<h2>Uploader Debug Test</h2>"
Response.Write "<p>TotalBytes: " & Request.TotalBytes & "</p>"
Response.Flush

' Simulating exactly what uploader does
Dim VarArrayBinRequest, StreamRequest, internalChunkSize, readBytes, tmpBinRequest
Dim errorMessage, UploadedFiles, FormElements, uploadedYet

internalChunkSize = 200000

' Initialize
Set UploadedFiles = Server.CreateObject("Scripting.Dictionary")
Set FormElements = Server.CreateObject("Scripting.Dictionary")
uploadedYet = False

' Create stream
Response.Write "<p>1. Creating stream...</p>"
Response.Flush
Set StreamRequest = Server.CreateObject("ADODB.Stream")
StreamRequest.Type = 1 ' adTypeBinary
StreamRequest.Open
Response.Write "<p>2. Stream created</p>"
Response.Flush

' Read binary data
Response.Write "<p>3. Reading binary data...</p>"
Response.Flush

On Error Resume Next
readBytes = internalChunkSize
VarArrayBinRequest = Request.BinaryRead(readBytes)
Response.Write "<p>4. First read: readBytes=" & readBytes & ", got " & LenB(VarArrayBinRequest) & " bytes</p>"
Response.Flush

VarArrayBinRequest = MidB(VarArrayBinRequest, 1, LenB(VarArrayBinRequest))
Response.Write "<p>5. After MidB</p>"
Response.Flush

Dim loopCount : loopCount = 0
Do Until readBytes < 1
    loopCount = loopCount + 1
    tmpBinRequest = Request.BinaryRead(readBytes)
    Response.Write "<p>Loop " & loopCount & ": readBytes=" & readBytes & "</p>"
    Response.Flush
    If loopCount > 5 Then Exit Do
    If readBytes > 0 Then
        VarArrayBinRequest = VarArrayBinRequest & MidB(tmpBinRequest, 1, LenB(tmpBinRequest))
    End If
Loop
Response.Write "<p>6. Loop complete, total: " & LenB(VarArrayBinRequest) & " bytes</p>"
Response.Flush

' Write to stream - this is WriteText with binary data (original uploader does this)
Response.Write "<p>7. Calling StreamRequest.WriteText...</p>"
Response.Flush
StreamRequest.WriteText VarArrayBinRequest
Response.Write "<p>8. WriteText complete</p>"
Response.Flush

Response.Write "<p>9. Calling StreamRequest.Flush...</p>"
Response.Flush
StreamRequest.Flush
Response.Write "<p>10. Stream Flush complete</p>"
Response.Flush

If Err.Number <> 0 Then 
    Response.Write "<p style='color:red'>Error: " & Err.Number & " - " & Err.Description & "</p>"
End If
On Error GoTo 0

' Now parse the multipart data
Response.Write "<hr><p><b>Parsing multipart data...</b></p>"
Response.Flush

' Token functions
Function String2Byte(sString)
    Dim i
    For i = 1 to Len(sString)
       String2Byte = String2Byte & ChrB(AscB(Mid(sString,i,1)))
    Next
End Function

Function FindToken(sToken, nStart)
    FindToken = InStrB(nStart, VarArrayBinRequest, sToken)
End Function

' Tokens
Dim tNewLine, tTerm, tFilename, tName, tContentDisp, tContentType, tDoubleQuotes
tNewLine = String2Byte(Chr(13))
tDoubleQuotes = String2Byte(Chr(34))
tTerm = String2Byte("--")
tFilename = String2Byte("filename=""")
tName = String2Byte("name=""")
tContentDisp = String2Byte("Content-Disposition")
tContentType = String2Byte("Content-Type:")

' Find first newline
Dim nCurPos, vDataSep, nDataBoundPos, nLastSepPos
nCurPos = FindToken(tNewLine, 1)
Response.Write "<p>11. First newline at: " & nCurPos & "</p>"
Response.Flush

If nCurPos <= 1 Then 
    Response.Write "<p style='color:red'>ERROR: No newline found!</p>"
    Response.End
End If

vDataSep = MidB(VarArrayBinRequest, 1, nCurPos-1)
Response.Write "<p>12. Separator length: " & LenB(vDataSep) & "</p>"
Response.Flush

nDataBoundPos = 1
nLastSepPos = FindToken(vDataSep & tTerm, 1)
Response.Write "<p>13. nDataBoundPos: " & nDataBoundPos & ", nLastSepPos: " & nLastSepPos & "</p>"
Response.Flush

' Parse loop
Dim parseLoopCount : parseLoopCount = 0
Response.Write "<p>14. Starting parse loop (nDataBoundPos <> nLastSepPos)...</p>"
Response.Flush

Do Until nDataBoundPos = nLastSepPos
    parseLoopCount = parseLoopCount + 1
    Response.Write "<p>  Parse loop " & parseLoopCount & ": nDataBoundPos=" & nDataBoundPos & ", nLastSepPos=" & nLastSepPos & "</p>"
    Response.Flush
    
    If parseLoopCount > 10 Then
        Response.Write "<p style='color:orange'>  SAFETY EXIT</p>"
        Exit Do
    End If
    
    ' Find Content-Disposition
    nCurPos = InStrB(nDataBoundPos, VarArrayBinRequest, tContentDisp)
    Response.Write "<p>  Content-Disposition at: " & nCurPos & "</p>"
    Response.Flush
    
    If nCurPos = 0 Then
        Response.Write "<p style='color:red'>  No Content-Disposition found!</p>"
        Exit Do
    End If
    nCurPos = nCurPos + LenB(tContentDisp)
    
    ' Find name="
    Dim namePos
    namePos = InStrB(nCurPos, VarArrayBinRequest, tName)
    Response.Write "<p>  name= at: " & namePos & "</p>"
    Response.Flush
    
    If namePos = 0 Then
        Response.Write "<p style='color:red'>  No name= found!</p>"
        Exit Do
    End If
    nCurPos = namePos + LenB(tName)
    
    ' Extract field name until "
    Dim nameEnd, sFieldName
    nameEnd = InStrB(nCurPos, VarArrayBinRequest, tDoubleQuotes)
    Response.Write "<p>  name end quote at: " & nameEnd & "</p>"
    Response.Flush
    
    If nameEnd > nCurPos Then
        ' Convert binary to string
        Dim nameBin, nameStr, j
        nameBin = MidB(VarArrayBinRequest, nCurPos, nameEnd - nCurPos)
        nameStr = ""
        For j = 1 To LenB(nameBin)
            nameStr = nameStr & Chr(AscB(MidB(nameBin, j, 1)))
        Next
        sFieldName = nameStr
        Response.Write "<p>  Field name: [" & sFieldName & "]</p>"
        Response.Flush
    End If
    
    ' Check if filename exists
    Dim nPosFile, nPosBound
    nPosFile = FindToken(tFilename, nCurPos)
    nPosBound = FindToken(vDataSep, nCurPos)
    Response.Write "<p>  filename= at: " & nPosFile & ", next boundary at: " & nPosBound & "</p>"
    Response.Flush
    
    If nPosFile <> 0 And nPosFile < nPosBound Then
        Response.Write "<p>  This is a FILE field</p>"
    Else
        Response.Write "<p>  This is a FORM field</p>"
    End If
    
    ' Advance to next boundary
    nDataBoundPos = FindToken(vDataSep, nCurPos)
    Response.Write "<p>  Next boundary at: " & nDataBoundPos & "</p>"
    Response.Flush
Loop

Response.Write "<hr><p><b>Parse complete! Iterations: " & parseLoopCount & "</b></p>"
Response.Flush

' Cleanup
StreamRequest.Close
Set StreamRequest = Nothing

Response.Write "<p>Test complete!</p>"
%>
