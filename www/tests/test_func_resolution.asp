<%@ Language="VBScript" %>
<%
Response.ContentType = "text/plain"
Response.Buffer = False
Response.Write "=== TEST FUNCTION RESOLUTION ===" & vbCrLf
Response.Flush

Function myGetFileType(filename)
    Response.Write "  Inside myGetFileType(" & filename & ")" & vbCrLf
    Response.Flush
    Dim dotPos
    dotPos = InStrRev(filename, ".", -1, 1)
    Response.Write "  InStrRev result: " & dotPos & vbCrLf
    Response.Flush
    myGetFileType = Right(filename, Len(filename) - dotPos)
    Response.Write "  Right result: " & myGetFileType & vbCrLf
    Response.Flush
End Function

Class TestFile
    Private nameOfFile
    
    Public Property Let FileName(fN)
        nameOfFile = fN
    End Property
    
    Public Property Get FileName()
        FileName = nameOfFile
    End Property
    
    Public Property Get FileType()
        Response.Write "FileType Get, calling myGetFileType..." & vbCrLf
        Response.Flush
        FileType = myGetFileType(FileName)
        Response.Write "FileType = " & FileType & vbCrLf
        Response.Flush
    End Property
End Class

Response.Write "Creating TestFile..." & vbCrLf
Response.Flush

Dim oFile
Set oFile = New TestFile
oFile.FileName = "test.txt"

Response.Write "Getting FileType..." & vbCrLf
Response.Flush

Dim ft
ft = oFile.FileType

Response.Write "Result: " & ft & vbCrLf
Response.Flush

Response.Write vbCrLf & "=== DONE ===" & vbCrLf
%>
