<%@ Language="VBScript" %>
<%
Response.ContentType = "text/plain"
Response.Buffer = False
Response.Write "=== TEST UPLOAD FILE TYPE LOGIC ===" & vbCrLf
Response.Flush

' Simulate what uploader does
Class TestUploadedFile
    Private nameOfFile
    
    Public Property Let FileName(fN)
        Response.Write "FileName Let: " & fN & vbCrLf
        Response.Flush
        nameOfFile = fN
        nameOfFile = SubstNoReg(nameOfFile, "\", "_")
        nameOfFile = SubstNoReg(nameOfFile, "/", "_")
        nameOfFile = SubstNoReg(nameOfFile, ":", "_")
        nameOfFile = SubstNoReg(nameOfFile, "*", "_")
        nameOfFile = SubstNoReg(nameOfFile, "?", "_")
        nameOfFile = SubstNoReg(nameOfFile, """", "_")
        nameOfFile = SubstNoReg(nameOfFile, "<", "_")
        nameOfFile = SubstNoReg(nameOfFile, ">", "_")
        nameOfFile = SubstNoReg(nameOfFile, "|", "_")
        Response.Write "After SubstNoReg: " & nameOfFile & vbCrLf
        Response.Flush
    End Property
    
    Public Property Get FileName()
        Response.Write "FileName Get called" & vbCrLf
        Response.Flush
        FileName = nameOfFile
    End Property
    
    Public Property Get FileType()
        Response.Write "FileType Get called, calling getFileType..." & vbCrLf
        Response.Flush
        FileType = getFileType(FileName)
        Response.Write "FileType result: " & FileType & vbCrLf
        Response.Flush
    End Property
End Class

Function SubstNoReg(initialStr, oldStr, newStr)
    Dim currentPos, oldStrPos, skip
    If IsNull(initialStr) Or Len(initialStr) = 0 Then
        SubstNoReg = ""
    ElseIf IsNull(oldStr) Or Len(oldStr) = 0 Then
        SubstNoReg = initialStr
    Else
        If IsNull(newStr) Then newStr = ""
        currentPos = 1
        oldStrPos = 0
        SubstNoReg = ""
        skip = Len(oldStr)
        Do While currentPos <= Len(initialStr)
            oldStrPos = InStr(currentPos, initialStr, oldStr)
            If oldStrPos = 0 Then
                SubstNoReg = SubstNoReg & Mid(initialStr, currentPos, Len(initialStr) - currentPos + 1)
                currentPos = Len(initialStr) + 1
            Else
                SubstNoReg = SubstNoReg & Mid(initialStr, currentPos, oldStrPos - currentPos) & newStr
                currentPos = oldStrPos + skip
            End If
        Loop
    End If
End Function

Function getFileType(filename)
    Response.Write "  getFileType(" & filename & ")" & vbCrLf
    Response.Flush
    Dim dotPos
    dotPos = InStrRev(filename, ".", -1, 1)
    Response.Write "  InStrRev result: " & dotPos & vbCrLf
    Response.Flush
    getFileType = Right(filename, Len(filename) - dotPos)
    Response.Write "  Right result: " & getFileType & vbCrLf
    Response.Flush
End Function

' Test
Dim oFile
Set oFile = New TestUploadedFile
Response.Write "Setting FileName..." & vbCrLf
Response.Flush
oFile.FileName = "test.txt"
Response.Write "Getting FileType..." & vbCrLf
Response.Flush
Dim ft
ft = oFile.FileType
Response.Write "FileType = " & ft & vbCrLf
Response.Flush

Response.Write vbCrLf & "=== DONE ===" & vbCrLf
%>
