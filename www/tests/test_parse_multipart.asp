<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include virtual="/asplite/asplite.asp"-->
<%
Response.ContentType = "text/plain"
Response.Write "=== MANUAL PARSING TEST ===" & vbCrLf

On Error Resume Next

' Get binary data
Dim readBytes, binData
readBytes = 5000
binData = Request.BinaryRead(readBytes)

Response.Write "Bytes read: " & readBytes & vbCrLf
Response.Write "LenB(binData): " & LenB(binData) & vbCrLf

' Create tokens like uploader does
Dim tNewLine, tDoubleQuotes, tFilename, tName, tContentDisp, tContentType
tNewLine = String2Byte(Chr(13))
tDoubleQuotes = String2Byte(Chr(34))
tFilename = String2Byte("filename=""")
tName = String2Byte("name=""")
tContentDisp = String2Byte("Content-Disposition")
tContentType = String2Byte("Content-Type:")

Response.Write vbCrLf & "Token lengths:" & vbCrLf
Response.Write "  tNewLine: " & LenB(tNewLine) & vbCrLf
Response.Write "  tDoubleQuotes: " & LenB(tDoubleQuotes) & vbCrLf
Response.Write "  tFilename: " & LenB(tFilename) & vbCrLf
Response.Write "  tName: " & LenB(tName) & vbCrLf
Response.Write "  tContentDisp: " & LenB(tContentDisp) & vbCrLf
Response.Write "  tContentType: " & LenB(tContentType) & vbCrLf

' Find first newline to get separator
Dim nCurPos
nCurPos = InstrB(1, binData, tNewLine)
Response.Write vbCrLf & "First newline position: " & nCurPos & vbCrLf

If nCurPos > 1 Then
    Dim vDataSep
    vDataSep = MidB(binData, 1, nCurPos - 1)
    Response.Write "Separator length: " & LenB(vDataSep) & vbCrLf
    Response.Write "Separator: " & vDataSep & vbCrLf
    
    ' Find Content-Disposition
    Dim posContentDisp
    posContentDisp = InstrB(1, binData, tContentDisp)
    Response.Write vbCrLf & "Content-Disposition position: " & posContentDisp & vbCrLf
    
    ' Find name="
    Dim posName
    posName = InstrB(posContentDisp, binData, tName)
    Response.Write "name="" position: " & posName & vbCrLf
    
    If posName > 0 Then
        ' Extract field name
        posName = posName + LenB(tName)
        Dim posEndName
        posEndName = InstrB(posName, binData, tDoubleQuotes)
        Response.Write "End of name position: " & posEndName & vbCrLf
        
        If posEndName > posName Then
            Dim fieldName
            fieldName = MidB(binData, posName, posEndName - posName)
            Response.Write "Field name: " & fieldName & vbCrLf
        End If
    End If
    
    ' Find filename="
    Dim posFilename
    posFilename = InstrB(posContentDisp, binData, tFilename)
    Response.Write vbCrLf & "filename="" position: " & posFilename & vbCrLf
    
    If posFilename > 0 Then
        posFilename = posFilename + LenB(tFilename)
        Dim posEndFilename
        posEndFilename = InstrB(posFilename, binData, tDoubleQuotes)
        Response.Write "End of filename position: " & posEndFilename & vbCrLf
        
        If posEndFilename > posFilename Then
            Dim fileName
            fileName = MidB(binData, posFilename, posEndFilename - posFilename)
            Response.Write "File name: " & fileName & vbCrLf
            Response.Write "File type: " & aspL.getFileType(fileName) & vbCrLf
        End If
    End If
End If

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
