<%@ Language="VBScript" %>
<%
Response.ContentType = "text/html"
Response.Buffer = False

If Request.TotalBytes < 1 Then
    Response.Write "<p>No data. POST multipart/form-data to this endpoint.</p>"
    Response.End
End If

Response.Write "<h2>Full Uploader Save Simulation</h2>"
Response.Write "<p>TotalBytes: " & Request.TotalBytes & "</p>"

'=== UPLOADER CLASS COPY ===
' Simulating exactly what cls_asplite_uploader does

Dim VarArrayBinRequest, StreamRequest, internalChunkSize
Dim UploadedFiles, FormElements, uploadedYet, overWriteFiles
Dim errorMessage

internalChunkSize = 200000
overWriteFiles = True
uploadedYet = False

Set UploadedFiles = Server.CreateObject("Scripting.Dictionary")
Set FormElements = Server.CreateObject("Scripting.Dictionary")

' Create stream
Set StreamRequest = Server.CreateObject("ADODB.Stream")
StreamRequest.Type = 1
StreamRequest.Open

Response.Write "<p>Stream created</p>"

' UploadedFile class
Class UploadedFile
    Public ContentType
    Public Start
    Public Length
    Public Path
    Private m_sFileName
    
    Public Property Let FileName(sFileName)
        m_sFileName = sFileName
    End Property
    
    Public Property Get FileName()
        FileName = m_sFileName
    End Property
    
    Public Property Get FileType()
        Dim pos
        pos = InStrRev(m_sFileName, ".")
        If pos > 0 Then
            FileType = LCase(Right(m_sFileName, Len(m_sFileName) - pos))
        Else
            FileType = ""
        End If
    End Property
    
    Public Property Get Size()
        Size = Length
    End Property
End Class

' Helper functions
Function String2Byte(sString)
    Dim i
    For i = 1 to Len(sString)
       String2Byte = String2Byte & ChrB(AscB(Mid(sString,i,1)))
    Next
End Function

Function FindToken(sToken, nStart)
    FindToken = InStrB(nStart, VarArrayBinRequest, sToken)
End Function

Function SkipToken(sToken, nStart)
    SkipToken = InStrB(nStart, VarArrayBinRequest, sToken)
    If SkipToken = 0 Then
        Response.Write "<p style='color:red'>SkipToken failed for token!</p>"
        SkipToken = nStart ' Avoid infinite loop
    Else
        SkipToken = SkipToken + LenB(sToken)
    End If
End Function

Function ExtractField(sToken, nStart)
    Dim nEnd
    nEnd = InStrB(nStart, VarArrayBinRequest, sToken)
    If nEnd = 0 Then
        Response.Write "<p style='color:red'>ExtractField: token not found!</p>"
        ExtractField = ""
        Exit Function
    End If
    
    Dim binData, strResult, j
    binData = MidB(VarArrayBinRequest, nStart, nEnd - nStart)
    strResult = ""
    For j = 1 To LenB(binData)
        strResult = strResult & Chr(AscB(MidB(binData, j, 1)))
    Next
    ExtractField = strResult
End Function

' === UPLOAD SUB ===
Response.Write "<hr><p><b>Running Upload...</b></p>"

Dim nCurPos, nDataBoundPos, nLastSepPos, nPosFile, nPosBound, sFieldName, osPathSep, auxStr, readBytes, tmpBinRequest
Dim vDataSep
Dim tNewLine, tDoubleQuotes, tTerm, tFilename, tName, tContentDisp, tContentType
tNewLine = String2Byte(Chr(13))
tDoubleQuotes = String2Byte(Chr(34))
tTerm = String2Byte("--")
tFilename = String2Byte("filename=""")
tName = String2Byte("name=""")
tContentDisp = String2Byte("Content-Disposition")
tContentType = String2Byte("Content-Type:")

uploadedYet = True

' Read binary data
readBytes = internalChunkSize
VarArrayBinRequest = Request.BinaryRead(readBytes)
VarArrayBinRequest = MidB(VarArrayBinRequest, 1, LenB(VarArrayBinRequest))
Do Until readBytes < 1
    tmpBinRequest = Request.BinaryRead(readBytes)
    If readBytes > 0 Then
        VarArrayBinRequest = VarArrayBinRequest & MidB(tmpBinRequest, 1, LenB(tmpBinRequest))
    End If
Loop

Response.Write "<p>Read " & LenB(VarArrayBinRequest) & " bytes</p>"

' Write to stream (uploader uses WriteText but we'll use Write for binary)
StreamRequest.Write VarArrayBinRequest
Response.Write "<p>Wrote to stream, size: " & StreamRequest.Size & "</p>"

' Find first newline
nCurPos = FindToken(tNewLine, 1)
Response.Write "<p>First newline at: " & nCurPos & "</p>"

If nCurPos <= 1 Then 
    Response.Write "<p style='color:red'>No data found!</p>"
    Response.End
End If

vDataSep = MidB(VarArrayBinRequest, 1, nCurPos-1)
nDataBoundPos = 1
nLastSepPos = FindToken(vDataSep & tTerm, 1)
Response.Write "<p>Separator found, last sep at: " & nLastSepPos & "</p>"

Dim parseLoop : parseLoop = 0
Do Until nDataBoundPos = nLastSepPos
    parseLoop = parseLoop + 1
    If parseLoop > 5 Then Exit Do
    
    Response.Write "<p>Parse iteration " & parseLoop & "</p>"
    
    nCurPos = SkipToken(tContentDisp, nDataBoundPos)
    nCurPos = SkipToken(tName, nCurPos)
    sFieldName = ExtractField(tDoubleQuotes, nCurPos)
    Response.Write "<p>  Field name: " & sFieldName & "</p>"
    
    nPosFile = FindToken(tFilename, nCurPos)
    nPosBound = FindToken(vDataSep, nCurPos)
    
    If nPosFile <> 0 And nPosFile < nPosBound Then
        Response.Write "<p>  FILE field detected</p>"
        
        Dim oUploadFile
        Set oUploadFile = New UploadedFile
        
        nCurPos = SkipToken(tFilename, nCurPos)
        auxStr = ExtractField(tDoubleQuotes, nCurPos)
        osPathSep = "\"
        If InStr(auxStr, osPathSep) = 0 Then osPathSep = "/"
        oUploadFile.FileName = Right(auxStr, Len(auxStr)-InStrRev(auxStr, osPathSep))
        Response.Write "<p>  Filename: " & oUploadFile.FileName & "</p>"
        
        If Len(oUploadFile.FileName) > 0 Then
            nCurPos = SkipToken(tContentType, nCurPos)
            auxStr = ExtractField(tNewLine, nCurPos)
            oUploadFile.ContentType = Right(auxStr, Len(auxStr)-InStrRev(auxStr, " "))
            nCurPos = FindToken(tNewLine, nCurPos) + 4
            
            oUploadFile.Start = nCurPos + 1
            oUploadFile.Length = FindToken(vDataSep, nCurPos) - 2 - nCurPos
            Response.Write "<p>  Start: " & oUploadFile.Start & ", Length: " & oUploadFile.Length & "</p>"
            
            If oUploadFile.Length > 0 Then
                Dim fileType
                fileType = oUploadFile.FileType
                Response.Write "<p>  FileType: " & fileType & "</p>"
                
                ' Add to dictionary (simplified - allow all types for testing)
                UploadedFiles.Add LCase(nCurPos & sFieldName), oUploadFile
                Response.Write "<p style='color:green'>  Added to UploadedFiles</p>"
            End If
        End If
    Else
        Response.Write "<p>  FORM field</p>"
    End If
    
    nDataBoundPos = FindToken(vDataSep, nCurPos)
Loop

Response.Write "<hr><p><b>Upload complete! Files: " & UploadedFiles.Count & "</b></p>"

' === SAVE SUB ===
Response.Write "<hr><p><b>Running Save...</b></p>"

Dim savePath
savePath = Server.MapPath("../uploads")
If Right(savePath, 1) <> "\" Then savePath = savePath & "\"
Response.Write "<p>Save path: " & savePath & "</p>"

Response.Write "<p>UploadedFiles.Count = " & UploadedFiles.Count & "</p>"
Response.Write "<p>UploadedFiles.Items type: " & TypeName(UploadedFiles.Items) & "</p>"

Dim itemsArr
itemsArr = UploadedFiles.Items
Response.Write "<p>Items array type: " & TypeName(itemsArr) & "</p>"
Response.Write "<p>Items array bounds: " & LBound(itemsArr) & " to " & UBound(itemsArr) & "</p>"

Response.Write "<p>Starting For Each loop over Items...</p>"
Dim fileItem, saveLoopCount
saveLoopCount = 0

For Each fileItem In UploadedFiles.Items
    saveLoopCount = saveLoopCount + 1
    Response.Write "<p>  Save loop " & saveLoopCount & "</p>"
    Response.Write "<p>  fileItem type: " & TypeName(fileItem) & "</p>"
    Response.Write "<p>  FileName: " & fileItem.FileName & "</p>"
    Response.Write "<p>  Start: " & fileItem.Start & ", Length: " & fileItem.Length & "</p>"
    
    Dim filePath
    filePath = savePath & fileItem.FileName
    Response.Write "<p>  Saving to: " & filePath & "</p>"
    
    ' Create output stream
    Dim streamFile
    Set streamFile = Server.CreateObject("ADODB.Stream")
    streamFile.Type = 1
    streamFile.Open
    Response.Write "<p>  Output stream created</p>"
    
    ' Position source stream
    Response.Write "<p>  Setting StreamRequest.Position to " & fileItem.Start & "</p>"
    StreamRequest.Position = fileItem.Start
    Response.Write "<p>  Position set, now: " & StreamRequest.Position & "</p>"
    
    ' Copy data
    Response.Write "<p>  Calling CopyTo with length " & fileItem.Length & "</p>"
    StreamRequest.CopyTo streamFile, fileItem.Length
    Response.Write "<p>  CopyTo complete, streamFile.Size = " & streamFile.Size & "</p>"
    
    ' Save file
    Response.Write "<p>  Saving file...</p>"
    streamFile.SaveToFile filePath, 2
    Response.Write "<p style='color:green'>  FILE SAVED!</p>"
    
    streamFile.Close
    Set streamFile = Nothing
Next

Response.Write "<hr><p><b>SAVE COMPLETE! Saved " & saveLoopCount & " file(s)</b></p>"

' Cleanup
StreamRequest.Close
Set StreamRequest = Nothing
%>
