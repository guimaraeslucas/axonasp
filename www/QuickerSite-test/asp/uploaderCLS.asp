
<%@ Language=VBScript%><%'  For examples, documentation, and your own free copy, go to:
'  http://www.freeaspupload.net
'  Note: You can copy and use this script for free and you can make changes
'  to the code, but you cannot remove the above comment.
'Changes:
'Aug 2, 2005: Add support for checkboxes and other input elements with multiple values
'Jan 6, 2009: Lars added ASP_CHUNK_SIZE
const DEFAULT_ASP_CHUNK_SIZE = 200000
Class Uploader
   Public UploadedFiles
   Public FormElements
   Private VarArrayBinRequest
   Private StreamRequest
   Private uploadedYet
   Private internalChunkSize
   Private Sub Class_Initialize()
      Set UploadedFiles = Server.CreateObject("Scripting.Dictionary")
      Set FormElements = Server.CreateObject("Scripting.Dictionary")
      Set StreamRequest = Server.CreateObject("ADODB.Stream")
      StreamRequest.Type = 2   ' adTypeText
      StreamRequest.Open
      uploadedYet = false
      internalChunkSize = DEFAULT_ASP_CHUNK_SIZE
   End Sub
   
   Private Sub Class_Terminate()
      If IsObject(UploadedFiles) Then
         UploadedFiles.RemoveAll()
         Set UploadedFiles = Nothing
      End If
      If IsObject(FormElements) Then
         FormElements.RemoveAll()
         Set FormElements = Nothing
      End If
      StreamRequest.Close
      Set StreamRequest = Nothing
   End Sub
   Public Property Get Form(sIndex)
      Form = ""
      If FormElements.Exists(LCase(sIndex)) Then Form = FormElements.Item(LCase(sIndex))
   End Property
   Public Property Get Files()
      Files = UploadedFiles.Items
   End Property
   
    Public Property Get Exists(sIndex)
            Exists = false
            If FormElements.Exists(LCase(sIndex)) Then Exists = true
    End Property
       
    Public Property Get FileExists(sIndex)
        FileExists = false
            if UploadedFiles.Exists(LCase(sIndex)) then FileExists = true
    End Property
       
    Public Property Get chunkSize()
      chunkSize = internalChunkSize
   End Property
   Public Property Let chunkSize(sz)
      internalChunkSize = sz
   End Property
   'Calls Upload to extract the data from the binary request and then saves the uploaded files
   Public Sub Save(path)
      Dim streamFile, fileItem
      if Right(path, 1) <> "\" then path = path & "\"
      if not uploadedYet then Upload
      For Each fileItem In UploadedFiles.Items
         Set streamFile = Server.CreateObject("ADODB.Stream")
         streamFile.Type = 1
         streamFile.Open
         StreamRequest.Position=fileItem.Start
         StreamRequest.CopyTo streamFile, fileItem.Length
         streamFile.SaveToFile GetFileName(path, fileItem.FileName), 2
         streamFile.close
         Set streamFile = Nothing
         fileItem.Path = path & fileItem.FileName
       Next
   
   
   
   End Sub
   
   public sub SaveOne(path, num, byref outFileName, byref outLocalFileName)
      Dim streamFile, fileItems, fileItem, fs
        set fs = Server.CreateObject("Scripting.FileSystemObject")
      if Right(path, 1) <> "\" then path = path & "\"
      if not uploadedYet then Upload
      if UploadedFiles.Count > 0 then
         fileItems = UploadedFiles.Items
         set fileItem = fileItems(num)
      
         outFileName = fileItem.FileName
         outLocalFileName = GetFileName(path, outFileName)
 
 
 
 
         Set streamFile = Server.CreateObject("ADODB.Stream")
         streamFile.Type = 1
         streamFile.Open
         StreamRequest.Position = fileItem.Start
         StreamRequest.CopyTo streamFile, fileItem.Length
         streamFile.SaveToFile path & outLocalFileName, 2
         streamFile.close
         Set streamFile = Nothing
         fileItem.Path = path & filename
      end if
   end sub
   Public Function SaveBinRequest(path) ' For debugging purposes
      StreamRequest.SaveToFile path & "\debugStream.bin", 2
   End Function
   Public Sub DumpData() 'only works if files are plain text
      Dim i, aKeys, f
      response.write "Form Items:<br>"
      aKeys = FormElements.Keys
      For i = 0 To FormElements.Count -1 ' Iterate the array
         response.write aKeys(i) & " = " & FormElements.Item(aKeys(i)) & "<BR>"
      Next
      response.write "Uploaded Files:<br>"
      For Each f In UploadedFiles.Items
         response.write "Name: " & f.FileName & "<br>"
         response.write "Type: " & f.ContentType & "<br>"
         response.write "Start: " & f.Start & "<br>"
         response.write "Size: " & f.Length & "<br>"
       Next
      End Sub
   Public Sub Upload()
      Dim nCurPos, nDataBoundPos, nLastSepPos
      Dim nPosFile, nPosBound
      Dim sFieldName, osPathSep, auxStr
      Dim readBytes, readLoop, tmpBinRequest
      
      'RFC1867 Tokens
      Dim vDataSep
      Dim tNewLine, tDoubleQuotes, tTerm, tFilename, tName, tContentDisp, tContentType
      tNewLine = String2Byte(Chr(13))
      tDoubleQuotes = String2Byte(Chr(34))
      tTerm = String2Byte("--")
      tFilename = String2Byte("filename=""")
      tName = String2Byte("name=""")
      tContentDisp = String2Byte("Content-Disposition")
      tContentType = String2Byte("Content-Type:")
      uploadedYet = true
      on error resume next
         readBytes = internalChunkSize
         VarArrayBinRequest = Request.BinaryRead(readBytes)
         VarArrayBinRequest = midb(VarArrayBinRequest, 1, lenb(VarArrayBinRequest))
         for readLoop = 0 to 300000
            tmpBinRequest = Request.BinaryRead(readBytes)
            if readBytes < 1 then exit for
            VarArrayBinRequest = VarArrayBinRequest & midb(tmpBinRequest, 1, lenb(tmpBinRequest))
         next
         if Err.Number <> 0 then
            response.write "<br><br><B>System reported this error:</B><p>"
            response.write Err.Description & "<p>"
            response.write "The most likely cause for this error is the incorrect setup of AspMaxRequestEntityAllowed in IIS MetaBase. Please see instructions in the <A HREF='http://www.freeaspupload.net/freeaspupload/requirements.asp'>requirements page of freeaspupload.net</A>.<p>"
            Exit Sub
         end if
      on error goto 0 'reset error handling
      nCurPos = FindToken(tNewLine,1) 'Note: nCurPos is 1-based (and so is InstrB, MidB, etc)
      If nCurPos <= 1  Then Exit Sub
      
      'vDataSep is a separator like -----------------------------21763138716045
      vDataSep = MidB(VarArrayBinRequest, 1, nCurPos-1)
      'Start of current separator
      nDataBoundPos = 1
      'Beginning of last line
      nLastSepPos = FindToken(vDataSep & tTerm, 1)
      Do Until nDataBoundPos = nLastSepPos
         
         nCurPos = SkipToken(tContentDisp, nDataBoundPos)
         nCurPos = SkipToken(tName, nCurPos)
         sFieldName = ExtractField(tDoubleQuotes, nCurPos)
         nPosFile = FindToken(tFilename, nCurPos)
         nPosBound = FindToken(vDataSep, nCurPos)
         
         If nPosFile <> 0 And  nPosFile < nPosBound Then
            Dim oUploadFile
            Set oUploadFile = New UploadedFile
            
            nCurPos = SkipToken(tFilename, nCurPos)
            auxStr = ExtractField(tDoubleQuotes, nCurPos)
                ' We are interested only in the name of the file, not the whole path
                ' Path separator is \ in windows, / in UNIX
                ' While IE seems to put the whole pathname in the stream, Mozilla seem to
                ' only put the actual file name, so UNIX paths may be rare. But not impossible.
                osPathSep = "\"
                if InStr(auxStr, osPathSep) = 0 then osPathSep = "/"
            oUploadFile.FileName = Right(auxStr, Len(auxStr)-InStrRev(auxStr, osPathSep))
            if (Len(oUploadFile.FileName) > 0) then 'File field not left empty
               nCurPos = SkipToken(tContentType, nCurPos)
               
                    auxStr = ExtractField(tNewLine, nCurPos)
                    ' NN on UNIX puts things like this in the stream:
                    '    ?? python py type=?? python application/x-python
               oUploadFile.ContentType = Right(auxStr, Len(auxStr)-InStrRev(auxStr, " "))
               nCurPos = FindToken(tNewLine, nCurPos) + 4 'skip empty line
               
               oUploadFile.Start = nCurPos+1
               oUploadFile.Length = FindToken(vDataSep, nCurPos) - 2 - nCurPos
               	   
If oUploadFile.Length > 0  Then 
select case lcase(getFileExtension(oUploadFile.FileName))
case "jpg","jpeg","jpe","jp2","jfif","gif","bmp","png","psd","eps","ico","tif","tiff","ai","raw","tga","mng","svg","doc","rtf","txt","wpd","wps","csv","xml","xsd","sql","pdf","xls","mdb","ppt","docx","xlsx","pptx","ppsx","artx","mp3","wma","mid","midi","mp4","mpg","mpeg","wav","ram","ra","avi","mov","flv","m4a","m4v","htm","html","css","swf","js","rar","zip","ogv","ogg","webm","tar","gz","eot","ttf","ics","woff"
UploadedFiles.Add LCase(sFieldName), oUploadFile
case else
'not allowed
end select
end if
   
   
   
   
            End If
         Else
            Dim nEndOfData
            nCurPos = FindToken(tNewLine, nCurPos) + 4 'skip empty line
            nEndOfData = FindToken(vDataSep, nCurPos) - 2
            If Not FormElements.Exists(LCase(sFieldName)) Then
               FormElements.Add LCase(sFieldName), Byte2String(MidB(VarArrayBinRequest, nCurPos, nEndOfData-nCurPos))
            else
                    FormElements.Item(LCase(sFieldName))= FormElements.Item(LCase(sFieldName)) & ", " & Byte2String(MidB(VarArrayBinRequest, nCurPos, nEndOfData-nCurPos))
                end if
         End If
         'Advance to next separator
         nDataBoundPos = FindToken(vDataSep, nCurPos)
      Loop
      StreamRequest.WriteText(VarArrayBinRequest)
  
  'QS change
application("doresize")=true
   End Sub
   Private Function SkipToken(sToken, nStart)
      SkipToken = InstrB(nStart, VarArrayBinRequest, sToken)
      If SkipToken = 0 then
         Response.write "Error in parsing uploaded binary request. The most likely cause for this error is the incorrect setup of AspMaxRequestEntityAllowed in IIS MetaBase. Please see instructions in the <A HREF='http://www.freeaspupload.net/freeaspupload/requirements.asp'>requirements page of freeaspupload.net</A>.<p>"
         Response.End
      end if
      SkipToken = SkipToken + LenB(sToken)
   End Function
   Private Function FindToken(sToken, nStart)
      FindToken = InstrB(nStart, VarArrayBinRequest, sToken)
   End Function
   Private Function ExtractField(sToken, nStart)
      Dim nEnd
      nEnd = InstrB(nStart, VarArrayBinRequest, sToken)
      If nEnd = 0 then
         Response.write "Error in parsing uploaded binary request."
         Response.End
      end if
      ExtractField = Byte2String(MidB(VarArrayBinRequest, nStart, nEnd-nStart))
   End Function
   'String to byte string conversion
   Private Function String2Byte(sString)
      Dim i
      For i = 1 to Len(sString)
         String2Byte = String2Byte & ChrB(AscB(Mid(sString,i,1)))
      Next
   End Function
   'Byte string to string conversion
   Private Function Byte2String(bsString)
      Dim i
      dim b1, b2, b3, b4
      Byte2String =""
      For i = 1 to LenB(bsString)
          if AscB(MidB(bsString,i,1)) < 128 then
              ' One byte
              Byte2String = Byte2String & ChrW(AscB(MidB(bsString,i,1)))
          elseif AscB(MidB(bsString,i,1)) < 224 then
              ' Two bytes
              b1 = AscB(MidB(bsString,i,1))
              b2 = AscB(MidB(bsString,i+1,1))
              Byte2String = Byte2String & ChrW((((b1 AND 28) / 4) * 256 + (b1 AND 3) * 64 + (b2 AND 63)))
              i = i + 1
          elseif AscB(MidB(bsString,i,1)) < 240 then
              ' Three bytes
              b1 = AscB(MidB(bsString,i,1))
              b2 = AscB(MidB(bsString,i+1,1))
              b3 = AscB(MidB(bsString,i+2,1))
              Byte2String = Byte2String & ChrW(((b1 AND 15) * 16 + (b2 AND 60)) * 256 + (b2 AND 3) * 64 + (b3 AND 63))
              i = i + 2
          else
              ' Four bytes
              b1 = AscB(MidB(bsString,i,1))
              b2 = AscB(MidB(bsString,i+1,1))
              b3 = AscB(MidB(bsString,i+2,1))
              b4 = AscB(MidB(bsString,i+3,1))
              ' Don't know how to handle this, I believe Microsoft doesn't support these characters so I replace them with a "^"
              'Byte2String = Byte2String & ChrW(((b1 AND 3) * 64 + (b2 AND 63)) & "," & (((b1 AND 28) / 4) * 256 + (b1 AND 3) * 64 + (b2 AND 63)))
              Byte2String = Byte2String & "^"
              i = i + 3
          end if
      Next
   End Function
End Class
Class UploadedFile
   Public ContentType
   Public Start
   Public Length
   Public Path
   Private nameOfFile
    ' Need to remove characters that are valid in UNIX, but not in Windows
    Public Property Let FileName(fN)
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
    End Property
    Public Property Get FileName()
        FileName = nameOfFile
    End Property
    'Public Property Get FileN()ame
End Class
' Does not depend on RegEx, which is not available on older VBScript
' Is not recursive, which means it will not run out of stack space
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
Function GetFileName(strSaveToPath, FileName)
'This function is used when saving a file to check there is not already a file with the same name so that you don't overwrite it.
'It adds numbers to the filename e.g. file.gif becomes file1.gif becomes file2.gif and so on.
'It keeps going until it returns a filename that does not exist.
'You could just create a filename from the ID field but that means writing the record - and it still might exist!
'N.B. Requires strSaveToPath variable to be available - and containing the path to save to
    Dim Counter
    Dim Flag
    Dim strTempFileName
    Dim FileExt
    Dim NewFullPath
    dim objFSO, p
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Counter = 0
    p = instrrev(FileName, ".")
    FileExt = mid(FileName, p+1)
    strTempFileName = left(FileName, p-1)
    NewFullPath = strSaveToPath & "\" & FileName
    Flag = False
   
    Do Until Flag = True
        If objFSO.FileExists(NewFullPath) = False Then
            Flag = True
            GetFileName = NewFullPath 'Mid(NewFullPath, InstrRev(NewFullPath, "\") + 1)
        Else
            Counter = Counter + 1
            NewFullPath = strSaveToPath & "\" & strTempFileName & Counter & "." & FileExt
        End If
    Loop
End Function
Function GetFileExtension(sFileName)
Dim Pos
Pos = instrrev(sFileName, ".")
If Pos>0 Then GetFileExtension = Mid(sFileName, Pos+1)
End Function%>
