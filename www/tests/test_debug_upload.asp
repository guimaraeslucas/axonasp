<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include virtual="/asplite/asplite.asp"-->
<%
Response.ContentType = "text/plain"
Response.Write "=== DEBUG UPLOAD PROCESS ===" & vbCrLf

On Error Resume Next

' Replicate what uploader does
Dim VarArrayBinRequest, StreamRequest, internalChunkSize, readBytes, tmpBinRequest
internalChunkSize = 200000

Set StreamRequest = Server.CreateObject("ADODB.Stream")
StreamRequest.Type = 2 ' adTypeText
StreamRequest.Open

readBytes = internalChunkSize
VarArrayBinRequest = Request.BinaryRead(readBytes)
VarArrayBinRequest = MidB(VarArrayBinRequest, 1, LenB(VarArrayBinRequest))

Do Until readBytes < 1
    tmpBinRequest = Request.BinaryRead(readBytes)
    If readBytes > 0 Then
        VarArrayBinRequest = VarArrayBinRequest & MidB(tmpBinRequest, 1, LenB(tmpBinRequest))
    End If
Loop

StreamRequest.WriteText(VarArrayBinRequest)
StreamRequest.Flush()

Response.Write "Binary LenB: " & LenB(VarArrayBinRequest) & vbCrLf
Response.Write "Stream Size: " & StreamRequest.Size & vbCrLf

' Tokens using ChrB approach (what uploader does)
Dim tNewLine, tDoubleQuotes, tTerm, tFilename, tName, tContentDisp, tContentType
tNewLine = ChrB(13)
tDoubleQuotes = ChrB(34)
tTerm = ChrB(45) & ChrB(45)
' Build tokens using loop
Dim i
tFilename = ""
For i = 1 to Len("filename=""")
   tFilename = tFilename & ChrB(AscB(Mid("filename=""",i,1)))
Next
tName = ""
For i = 1 to Len("name=""")
   tName = tName & ChrB(AscB(Mid("name=""",i,1)))
Next
tContentDisp = ""
For i = 1 to Len("Content-Disposition")
   tContentDisp = tContentDisp & ChrB(AscB(Mid("Content-Disposition",i,1)))
Next
tContentType = ""
For i = 1 to Len("Content-Type:")
   tContentType = tContentType & ChrB(AscB(Mid("Content-Type:",i,1)))
Next

' Find first newline to get separator
Dim nCurPos, vDataSep, nDataBoundPos, nLastSepPos
nCurPos = InstrB(1, VarArrayBinRequest, tNewLine)
Response.Write vbCrLf & "First newline at: " & nCurPos & vbCrLf

If nCurPos > 1 Then

    vDataSep = MidB(VarArrayBinRequest, 1, nCurPos - 1)
    Response.Write "Separator LenB: " & LenB(vDataSep) & vbCrLf
    
    nDataBoundPos = 1
    nLastSepPos = InstrB(1, VarArrayBinRequest, vDataSep & tTerm)
    Response.Write "Last separator at: " & nLastSepPos & vbCrLf
    
    ' Find Content-Disposition
    nCurPos = InstrB(nDataBoundPos, VarArrayBinRequest, tContentDisp)
    Response.Write vbCrLf & "Content-Disposition at: " & nCurPos & vbCrLf
    
    If nCurPos > 0 Then
        nCurPos = nCurPos + LenB(tContentDisp)
        
        ' Find name
        Dim nNamePos
        nNamePos = InstrB(nCurPos, VarArrayBinRequest, tName)
        Response.Write "name= at: " & nNamePos & vbCrLf
        
        If nNamePos > 0 Then
            nCurPos = nNamePos + LenB(tName)
            
            ' Now check for filename
            Dim nPosFile, nPosBound
            nPosFile = InstrB(nCurPos, VarArrayBinRequest, tFilename)
            nPosBound = InstrB(nCurPos, VarArrayBinRequest, vDataSep)
            Response.Write vbCrLf & "filename= at: " & nPosFile & vbCrLf
            Response.Write "Next boundary at: " & nPosBound & vbCrLf
            
            If nPosFile <> 0 And nPosFile < nPosBound Then
                Response.Write "THIS IS A FILE UPLOAD!" & vbCrLf
                
                ' Move to after filename="
                nCurPos = nPosFile + LenB(tFilename)
                
                ' Find closing quote
                Dim nEnd
                nEnd = InstrB(nCurPos, VarArrayBinRequest, tDoubleQuotes)
                Response.Write "Filename ends at: " & nEnd & vbCrLf
                
                ' Extract filename using byte extraction
                Dim objStream, strTmp, fileNameStr, justFileName
                
                StreamRequest.Position = 0
                Set objStream = Server.CreateObject("ADODB.Stream")
                objStream.Charset = "utf-8"
                objStream.Mode = 3
                objStream.Type = 1
                objStream.Open
                ' InstrB returns 1-based, Position is 0-based, so subtract 1
                StreamRequest.Position = nCurPos - 1
                StreamRequest.CopyTo objStream, nEnd - nCurPos
                objStream.Flush
                objStream.Position = 0
                objStream.Type = 2
                fileNameStr = objStream.ReadText
                objStream.Close
                Set objStream = Nothing
                
                Response.Write "Filename extracted: [" & fileNameStr & "]" & vbCrLf
                
                ' Get just the filename
                Dim osPathSep
                osPathSep = "\"
                If InStr(fileNameStr, osPathSep) = 0 Then osPathSep = "/"
                justFileName = Right(fileNameStr, Len(fileNameStr) - InStrRev(fileNameStr, osPathSep))
                Response.Write "Just filename: [" & justFileName & "]" & vbCrLf
                Response.Write "Len(justFileName): " & Len(justFileName) & vbCrLf
                
                If Len(justFileName) > 0 Then
                    Response.Write "Filename has length > 0" & vbCrLf
                    
                    ' Find Content-Type
                    Dim nCTPos
                    nCTPos = InstrB(nCurPos, VarArrayBinRequest, tContentType)
                    Response.Write "Content-Type at: " & nCTPos & vbCrLf
                    
                    If nCTPos > 0 Then
                        nCurPos = nCTPos + LenB(tContentType)
                        nEnd = InstrB(nCurPos, VarArrayBinRequest, tNewLine)
                        
                        ' Skip empty line - this is +4 (CRLF CRLF)
                        nCurPos = nEnd + 4
                        Response.Write "File content starts at: " & nCurPos & vbCrLf
                        
                        ' Find end of file
                        Dim fileStart, fileLength
                        fileStart = nCurPos + 1
                        fileLength = InstrB(nCurPos, VarArrayBinRequest, vDataSep) - 2 - nCurPos
                        Response.Write "File start (for stream): " & fileStart & vbCrLf
                        Response.Write "File length: " & fileLength & vbCrLf
                        
                        If fileLength > 0 Then
                            Response.Write "FILE HAS CONTENT!" & vbCrLf
                            
                            ' Test the FileType
                            Dim ft
                            ft = aspL.getFileType(justFileName)
                            Response.Write "FileType: [" & ft & "]" & vbCrLf
                            Response.Write "LCase(FileType): [" & LCase(ft) & "]" & vbCrLf
                            
                            ' Check if allowed
                            Select Case LCase(ft)
                                Case "jpg","jpeg","jpe","jp2","jfif","gif","bmp","png","psd","eps","ico","tif","tiff","ai","raw","tga","mng","svg","doc","rtf","txt","wpd","wps","csv","xml","xsd","sql","pdf","xls","mdb","ppt","docx","xlsx","pptx","ppsx","artx","mp3","wma","mid","midi","mp4","mpg","mpeg","wav","ram","ra","avi","mov","flv","m4a","m4v","htm","html","css","swf","js","rar","zip","ogv","ogg","webm","tar","gz","eot","ttf","ics","woff","cod","msg","odt"
                                    Response.Write "FILE TYPE IS ALLOWED - WOULD BE ADDED!" & vbCrLf
                                Case Else
                                    Response.Write "FILE TYPE NOT ALLOWED: [" & LCase(ft) & "]" & vbCrLf
                            End Select
                        Else
                            Response.Write "File has no content (length=" & fileLength & ")" & vbCrLf
                        End If
                    End If
                Else
                    Response.Write "Filename is empty" & vbCrLf
                End If
            Else
                Response.Write "Not a file upload (nPosFile=" & nPosFile & ", nPosBound=" & nPosBound & ")" & vbCrLf
            End If
        Else
            Response.Write "Cannot find name=" & vbCrLf
        End If
    Else
        Response.Write "Cannot find Content-Disposition" & vbCrLf
    End If
Else
    Response.Write "Cannot find newline" & vbCrLf
End If

If Err.Number <> 0 Then
    Response.Write vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf
    Err.Clear
End If

StreamRequest.Close
Set StreamRequest = Nothing

Response.Write vbCrLf & "=== DONE ===" & vbCrLf
%>
