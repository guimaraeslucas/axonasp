<%@ Language="VBScript" %>
<%
' Minimal upload test - no aspLite
Response.ContentType = "text/plain"
Response.Write "=== MINIMAL UPLOAD TEST ===" & vbCrLf

If Request.TotalBytes < 1 Then
    Response.Write "No data. POST multipart to this endpoint." & vbCrLf
    Response.End
End If

' Read binary data
Dim binData, readBytes
readBytes = Request.TotalBytes
binData = Request.BinaryRead(readBytes)
Response.Write "Read " & LenB(binData) & " bytes" & vbCrLf

' Create stream for parsing
Dim StreamRequest
Set StreamRequest = Server.CreateObject("ADODB.Stream")
StreamRequest.Type = 2
StreamRequest.Open
StreamRequest.WriteText binData
StreamRequest.Flush
Response.Write "Stream size: " & StreamRequest.Size & vbCrLf

' RFC1867 Tokens
Dim tNewLine, tDoubleQuotes, tFilename, tName
tNewLine = ChrB(13)
tDoubleQuotes = ChrB(34)
tFilename = ChrB(102) & ChrB(105) & ChrB(108) & ChrB(101) & ChrB(110) & ChrB(97) & ChrB(109) & ChrB(101) & ChrB(61) & ChrB(34)
tName = ChrB(110) & ChrB(97) & ChrB(109) & ChrB(101) & ChrB(61) & ChrB(34)

' Find first CRLF
Dim nCurPos
nCurPos = InstrB(1, binData, tNewLine)
Response.Write "First CR at: " & nCurPos & vbCrLf

If nCurPos <= 1 Then
    Response.Write "ERROR: Could not find first CR" & vbCrLf
    Response.End
End If

' Extract boundary
Dim vDataSep
vDataSep = MidB(binData, 1, nCurPos - 1)
Response.Write "Boundary length: " & LenB(vDataSep) & vbCrLf

' Find name=""
Dim nPosName
nPosName = InstrB(1, binData, tName)
Response.Write "name="" at: " & nPosName & vbCrLf

If nPosName > 0 Then
    Dim nEndName
    nEndName = InstrB(nPosName + LenB(tName), binData, tDoubleQuotes)
    Response.Write "End of name at: " & nEndName & vbCrLf
    
    If nEndName > 0 Then
        Dim nameStart, nameLen
        nameStart = nPosName + LenB(tName) - 1
        nameLen = nEndName - (nPosName + LenB(tName))
        
        Dim objStream
        Set objStream = Server.CreateObject("ADODB.Stream")
        objStream.Charset = "utf-8"
        objStream.Mode = 3
        objStream.Type = 1
        objStream.Open
        
        StreamRequest.Position = nameStart
        StreamRequest.CopyTo objStream, nameLen
        
        objStream.Position = 0
        objStream.Type = 2
        
        Dim fieldName
        fieldName = objStream.ReadText
        Response.Write "Field name: [" & fieldName & "]" & vbCrLf
        
        objStream.Close
        Set objStream = Nothing
    End If
End If

' Find filename=""
Dim nPosFilename
nPosFilename = InstrB(1, binData, tFilename)
Response.Write "filename="" at: " & nPosFilename & vbCrLf

If nPosFilename > 0 Then
    Dim nEndFilename
    nEndFilename = InstrB(nPosFilename + LenB(tFilename), binData, tDoubleQuotes)
    Response.Write "End of filename at: " & nEndFilename & vbCrLf
    
    If nEndFilename > 0 Then
        Dim fnStart, fnLen
        fnStart = nPosFilename + LenB(tFilename) - 1
        fnLen = nEndFilename - (nPosFilename + LenB(tFilename))
        
        Dim objStream2
        Set objStream2 = Server.CreateObject("ADODB.Stream")
        objStream2.Charset = "utf-8"
        objStream2.Mode = 3
        objStream2.Type = 1
        objStream2.Open
        
        StreamRequest.Position = fnStart
        StreamRequest.CopyTo objStream2, fnLen
        
        objStream2.Position = 0
        objStream2.Type = 2
        
        Dim fileName
        fileName = objStream2.ReadText
        Response.Write "Filename: [" & fileName & "]" & vbCrLf
        
        objStream2.Close
        Set objStream2 = Nothing
    End If
End If

StreamRequest.Close
Set StreamRequest = Nothing

Response.Write vbCrLf & "=== SUCCESS ===" & vbCrLf
%>
