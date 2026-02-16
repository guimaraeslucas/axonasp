<%
'  DEBUG VERSION OF UPLOADER - For testing only

Class cls_asplite_uploader_debug

	Public UploadedFiles, FormElements, errorMessage, overWriteFiles
	Private VarArrayBinRequest, StreamRequest, uploadedYet, internalChunkSize

	Private Sub Class_Initialize()
		Response.Write "<p>DEBUG: Class_Initialize starting...</p>" : Response.Flush
		Set UploadedFiles = aspL.dict
		Response.Write "<p>DEBUG: UploadedFiles created</p>" : Response.Flush
		Set FormElements = aspL.dict
		Response.Write "<p>DEBUG: FormElements created</p>" : Response.Flush
		Set StreamRequest = Server.CreateObject("ADODB.Stream")
		Response.Write "<p>DEBUG: StreamRequest created</p>" : Response.Flush
		StreamRequest.Type = 2 'adTypeText
		StreamRequest.Open
		Response.Write "<p>DEBUG: StreamRequest opened, Type=" & StreamRequest.Type & "</p>" : Response.Flush
		uploadedYet = false
		overWriteFiles = false
		internalChunkSize = 200000
		Response.Write "<p>DEBUG: Class_Initialize complete</p>" : Response.Flush
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

	Public Sub Upload()
		Response.Write "<p>DEBUG: Upload() starting...</p>" : Response.Flush
		
		Dim nCurPos, nDataBoundPos, nLastSepPos, nPosFile, nPosBound, sFieldName, osPathSep, auxStr, readBytes, readLoop, tmpBinRequest
		
		'RFC1867 Tokens
		Dim vDataSep
		Dim tNewLine, tDoubleQuotes, tTerm, tFilename, tName, tContentDisp, tContentType
		
		Response.Write "<p>DEBUG: Creating tokens...</p>" : Response.Flush
		tNewLine = String2Byte(Chr(13))
		Response.Write "<p>DEBUG: tNewLine created, len=" & LenB(tNewLine) & "</p>" : Response.Flush
		tDoubleQuotes = String2Byte(Chr(34))
		tTerm = String2Byte("--")
		tFilename = String2Byte("filename=""")
		tName = String2Byte("name=""")
		tContentDisp = String2Byte("Content-Disposition")
		tContentType = String2Byte("Content-Type:")
		Response.Write "<p>DEBUG: All tokens created</p>" : Response.Flush

		uploadedYet = true

		on error resume next
			' Copy binary request to a byte array
			Response.Write "<p>DEBUG: Reading binary data...</p>" : Response.Flush
			readBytes = internalChunkSize
			VarArrayBinRequest = Request.BinaryRead(readBytes)
			Response.Write "<p>DEBUG: First read: readBytes=" & readBytes & ", got=" & LenB(VarArrayBinRequest) & "</p>" : Response.Flush
			
			VarArrayBinRequest = midb(VarArrayBinRequest, 1, lenb(VarArrayBinRequest))
			Response.Write "<p>DEBUG: After MidB</p>" : Response.Flush
			
			Dim loopCount : loopCount = 0
			Do Until readBytes < 1
				loopCount = loopCount + 1
				Response.Write "<p>DEBUG: Loop " & loopCount & ", readBytes=" & readBytes & "</p>" : Response.Flush
				tmpBinRequest = Request.BinaryRead(readBytes)
				Response.Write "<p>DEBUG: After BinaryRead: readBytes=" & readBytes & "</p>" : Response.Flush
				if readBytes > 0 then
					VarArrayBinRequest = VarArrayBinRequest & midb(tmpBinRequest, 1, lenb(tmpBinRequest))
				end if
			Loop
			Response.Write "<p>DEBUG: Loop complete, total=" & LenB(VarArrayBinRequest) & "</p>" : Response.Flush
			
			Response.Write "<p>DEBUG: Calling WriteText...</p>" : Response.Flush
			StreamRequest.WriteText(VarArrayBinRequest)
			Response.Write "<p>DEBUG: WriteText done</p>" : Response.Flush
			
			Response.Write "<p>DEBUG: Calling Flush...</p>" : Response.Flush
			StreamRequest.Flush()
			Response.Write "<p>DEBUG: Flush done</p>" : Response.Flush
			
			if Err.Number <> 0 then 
				errorMessage="<p>System reported an error:</p>"
				errorMessage=errorMessage & "<p>" & Err.Description  & "</p>"
				errorMessage=errorMessage & "<p>The most likely cause for this error is the incorrect setup of AspMaxRequestEntityAllowed in IIS MetaBase.</p>"
				Response.Write "<p>DEBUG: Error: " & Err.Description & "</p>" : Response.Flush
				Exit Sub
			end if
		on error goto 0

		Response.Write "<p>DEBUG: Finding tokens in data...</p>" : Response.Flush
		nCurPos = FindToken(tNewLine,1)
		Response.Write "<p>DEBUG: nCurPos=" & nCurPos & "</p>" : Response.Flush

		If nCurPos <= 1 Then 
			Response.Write "<p>DEBUG: Early exit, no data</p>" : Response.Flush
			Exit Sub
		End If
		 
		vDataSep = MidB(VarArrayBinRequest, 1, nCurPos-1)
		nDataBoundPos = 1
		nLastSepPos = FindToken(vDataSep & tTerm, 1)
		Response.Write "<p>DEBUG: nLastSepPos=" & nLastSepPos & "</p>" : Response.Flush
		
		Dim parseLoop : parseLoop = 0
		Do Until nDataBoundPos = nLastSepPos
			parseLoop = parseLoop + 1
			Response.Write "<p>DEBUG: Parse loop " & parseLoop & "</p>" : Response.Flush
			If parseLoop > 10 Then
				Response.Write "<p>DEBUG: PARSE SAFETY EXIT</p>" : Response.Flush
				Exit Do
			End If
			
			nCurPos = SkipToken(tContentDisp, nDataBoundPos)
			nCurPos = SkipToken(tName, nCurPos)
			sFieldName = ExtractField(tDoubleQuotes, nCurPos)
			Response.Write "<p>DEBUG: Field: " & sFieldName & "</p>" : Response.Flush

			nPosFile = FindToken(tFilename, nCurPos)
			nPosBound = FindToken(vDataSep, nCurPos)
			
			If nPosFile <> 0 And nPosFile < nPosBound Then
				Response.Write "<p>DEBUG: FILE field</p>" : Response.Flush
				
				Dim oUploadFile
				Set oUploadFile = New UploadedFileDebug
				
				nCurPos = SkipToken(tFilename, nCurPos)
				auxStr = ExtractField(tDoubleQuotes, nCurPos)
                osPathSep = "\"
                if InStr(auxStr, osPathSep) = 0 then osPathSep = "/"
				oUploadFile.FileName = Right(auxStr, Len(auxStr)-InStrRev(auxStr, osPathSep))
				Response.Write "<p>DEBUG: FileName: " & oUploadFile.FileName & "</p>" : Response.Flush

				if (Len(oUploadFile.FileName) > 0) then
					nCurPos = SkipToken(tContentType, nCurPos)
					
                    auxStr = ExtractField(tNewLine, nCurPos)
					oUploadFile.ContentType = Right(auxStr, Len(auxStr)-InStrRev(auxStr, " "))
					nCurPos = FindToken(tNewLine, nCurPos) + 4
					
					oUploadFile.Start = nCurPos+1
					oUploadFile.Length = FindToken(vDataSep, nCurPos) - 2 - nCurPos
					Response.Write "<p>DEBUG: Start=" & oUploadFile.Start & ", Length=" & oUploadFile.Length & "</p>" : Response.Flush
					
					If oUploadFile.Length > 0 Then 
						' Allow all file types for testing
						UploadedFiles.Add LCase(nCurPos & sFieldName), oUploadFile
						Response.Write "<p style='color:green'>DEBUG: Added to UploadedFiles</p>" : Response.Flush
					end if
				End If
			Else
				Response.Write "<p>DEBUG: FORM field</p>" : Response.Flush
				Dim nEndOfData, fieldValueUniStr
				nCurPos = FindToken(tNewLine, nCurPos) + 4
				nEndOfData = FindToken(vDataSep, nCurPos) - 2
				fieldValueuniStr = ConvertUtf8BytesToString(nCurPos, nEndOfData-nCurPos)
				If Not FormElements.Exists(LCase(sFieldName)) Then 
					FormElements.Add LCase(nCurPos & sFieldName), fieldValueuniStr
                else
                    FormElements.Item(LCase(sFieldName))= FormElements.Item(LCase(sFieldName)) & ", " & fieldValueuniStr
                end if 
			End If

			nDataBoundPos = FindToken(vDataSep, nCurPos)
		Loop
		
		Response.Write "<p>DEBUG: Upload() complete! Files=" & UploadedFiles.Count & "</p>" : Response.Flush
	End Sub

	Private Function SkipToken(sToken, nStart)
		SkipToken = InstrB(nStart, VarArrayBinRequest, sToken)
		If SkipToken = 0 then		
			Response.Write "<p style='color:red'>DEBUG: SkipToken failed!</p>" : Response.Flush
			SkipToken = nStart
		else
			SkipToken = SkipToken + LenB(sToken)
		end if
	End Function

	Private Function FindToken(sToken, nStart)
		FindToken = InstrB(nStart, VarArrayBinRequest, sToken)
	End Function

	Private Function ExtractField(sToken, nStart)
		Dim nEnd
		nEnd = InstrB(nStart, VarArrayBinRequest, sToken)
		If nEnd = 0 then
			Response.Write "<p style='color:red'>DEBUG: ExtractField failed!</p>" : Response.Flush
			ExtractField = ""
			Exit Function
		end if
		Dim binData, strRes, j
		binData = MidB(VarArrayBinRequest, nStart, nEnd - nStart)
		strRes = ""
		For j = 1 To LenB(binData)
			strRes = strRes & Chr(AscB(MidB(binData, j, 1)))
		Next
		ExtractField = strRes
	End Function

	Private Function String2Byte(sString)
		Dim i
		For i = 1 to Len(sString)
		   String2Byte = String2Byte & ChrB(AscB(Mid(sString,i,1)))
		Next
	End Function

	Private Function ConvertUtf8BytesToString(start, length)	
		StreamRequest.Position = 0
	
	    Dim objStream
	    Dim strTmp
	    
	    Set objStream = Server.CreateObject("ADODB.Stream")
	    objStream.Charset = "utf-8"
	    objStream.Type = 2
	    objStream.Open  
	    StreamRequest.Position = start
	    StreamRequest.CopyTo objStream, length  
	    objStream.Position = 0  
	    strTmp = objStream.ReadText()
	    objStream.Close
	    Set objStream = Nothing
	    ConvertUtf8BytesToString = strTmp
	End Function

End Class

Class UploadedFileDebug
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
%>
