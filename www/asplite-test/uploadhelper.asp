<%
Const MAX_UPLOAD_SIZE        = 500000000 '50 MB
Const MSG_NO_DATA            = "Nothing to upload."
Const MSG_EXCEEDED_MAX_SIZE  = "You exceeded the maximum upload size."
Const MSG_BAD_REQUEST_METHOD = "Bad request method. Use the POST method."
Const MSG_BAD_ENCTYPE        = "Bad encoding type. Use a ""multipart/form-data"" enctype."
Const MSG_ZERO_LENGTH        = "Zero length request."

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ':::                                                                 :::
    ':::  The file was modified by Shekhar Tarare (shekhartarare.com)    :::
    ':::  on 20/06/2023. I have fixed issues with the file by making     :::
    ':::  some changes to the code. It's working as of now. Feel free to ::: 
    ':::  modify according to your needs.                                :::                                
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Class UploadHelper
    Private m_Request
    Private m_Files
    Private m_Error

    Public Property Get GetError
        GetError = m_Error
    End Property

    Public Property Get FileCount
        FileCount = m_Files.Count
    End Property

    Public Function File(index)
        If m_Files.Exists(index) Then
           Set File = m_Files(index)
        Else  
           Set File = Nothing
        End If
    End Function

    Public Default Property Get Item(strName)
        If m_Request.Exists(strName) Then
            Item = m_Request(strName)
        Else  
            Item = ""
        End If
    End Property

    Private Sub Class_Initialize
        Dim iBytesCount, strBinData

        'first of all, get amount of uploaded bytes:
        iBytesCount = Request.TotalBytes

        'abort if nothing there:
        If iBytesCount = 0 Then
            m_Error = MSG_NO_DATA
            Exit Sub
        End If

        'abort if exceeded maximum upload size:
        If iBytesCount > MAX_UPLOAD_SIZE Then
            m_Error = MSG_EXCEEDED_MAX_SIZE
            Exit Sub
        End If

        If Request.ServerVariables("REQUEST_METHOD") = "POST" Then 
            Dim CT, PosB, Boundary, PosE
            CT = Request.ServerVariables("HTTP_Content_Type")
            If LCase(Left(CT, 19)) = "multipart/form-data" Then
                PosB = InStr(LCase(CT), "boundary=") 
                If PosB > 0 Then Boundary = Mid(CT, PosB + 9) 
                PosB = InStr(LCase(CT), "boundary=") 
                If PosB > 0 Then 
                    PosB = InStr(Boundary, ",")
                    If PosB > 0 Then Boundary = Left(Boundary, PosB - 1)
                End If
                If iBytesCount > 0 And Boundary <> "" Then 
                    Boundary = "--" & Boundary
                    Dim Head, Binary
                    Binary = Request.BinaryRead(iBytesCount) 

                    'create private collections:
                    Set m_Request = Server.CreateObject("Scripting.Dictionary")
                    Set m_Files = Server.CreateObject("Scripting.Dictionary")

                    Call ParseRequest(Binary, Boundary)
                    Binary = Empty 
                Else
                    m_Error = MSG_ZERO_LENGTH
                    Exit Sub
                End If
            Else
                m_Error = MSG_BAD_ENCTYPE
                Exit Sub
            End If
        Else
            m_Error = MSG_BAD_REQUEST_METHOD
            Exit Sub
        End If
    End Sub

    Private Sub Class_Terminate
        Dim fileName
        If IsObject(m_Request) Then
            m_Request.RemoveAll
            Set m_Request = Nothing
        End If
        If IsObject(m_Files) Then
            For Each fileName In m_Files.Keys
                Set m_Files(fileName)=Nothing
            Next
            m_Files.RemoveAll
            Set m_Files = Nothing
        End If
    End Sub

    Private Sub ParseRequest(Binary, Boundary)
        Dim PosOpenBoundary, PosCloseBoundary, PosEndOfHeader, isLastBoundary
        Boundary = StringToBinary(Boundary)
        PosOpenBoundary = InStrB(Binary, Boundary)
        PosCloseBoundary = InStrB(PosOpenBoundary + LenB(Boundary), Binary, Boundary, 0)
        Dim HeaderContent, FieldContent, bFieldContent
        Dim Content_Disposition, FormFieldName, SourceFileName, Content_Type
        Dim TwoCharsAfterEndBoundary, n : n = 0
        Do While (PosOpenBoundary > 0 And PosCloseBoundary > 0 And Not isLastBoundary)
            PosEndOfHeader = InStrB(PosOpenBoundary + Len(Boundary), Binary, StringToBinary(vbCrLf + vbCrLf))
            HeaderContent = MidB(Binary, PosOpenBoundary + LenB(Boundary) + 2, PosEndOfHeader - PosOpenBoundary - LenB(Boundary) - 2)
            bFieldContent = MidB(Binary, (PosEndOfHeader + 4), PosCloseBoundary - (PosEndOfHeader + 4) - 2)
            GetHeadFields BinaryToString(HeaderContent), Content_Disposition, FormFieldName, SourceFileName, Content_Type
            Set objFileData = New FileData
            objFileData.FileName = SourceFileName
            objFileData.ContentType = Content_Type
            objFileData.Contents = bFieldContent
            objFileData.FormFieldName = FormFieldName
            objFileData.ContentDisposition = Content_Disposition
            Set m_Files(n) = objFileData
            Set objFileData = Nothing
            TwoCharsAfterEndBoundary = BinaryToString(MidB(Binary, PosCloseBoundary + LenB(Boundary), 2))
            isLastBoundary = TwoCharsAfterEndBoundary = "--"
            If Not isLastBoundary Then 
                PosOpenBoundary = PosCloseBoundary
                PosCloseBoundary = InStrB(PosOpenBoundary + LenB(Boundary), Binary, Boundary)
            End If
            n = n + 1
        Loop
    End Sub

    Private Function GetHeadFields(ByVal Head, Content_Disposition, Name, FileName, Content_Type)
        Content_Disposition = LTrim(SeparateField(Head, "content-disposition:", ";"))
        Name = (SeparateField(Head, "name=", ";")) 
        If Left(Name, 1) = """" Then Name = Mid(Name, 2, Len(Name) - 2)
        FileName = (SeparateField(Head, "filename=", ";")) 
        If Left(FileName, 1) = """" Then FileName = Mid(FileName, 2, Len(FileName) - 2)
        Content_Type = LTrim(SeparateField(Head, "content-type:", ";"))
    End Function

    Private Function SeparateField(From, ByVal sStart, ByVal sEnd)
        Dim PosB, PosE, sFrom
        sFrom = LCase(From)
        PosB = InStr(sFrom, sStart)
        If PosB > 0 Then
            PosB = PosB + Len(sStart)
            PosE = InStr(PosB, sFrom, sEnd)
            If PosE = 0 Then PosE = InStr(PosB, sFrom, vbCrLf)
            If PosE = 0 Then PosE = Len(sFrom) + 1
            SeparateField = Mid(From, PosB, PosE - PosB)
        Else
            SeparateField = Empty
        End If
    End Function

    Private Function BinaryToString(Binary)
        dim cl1, cl2, cl3, pl1, pl2, pl3
        Dim L
        cl1 = 1
        cl2 = 1
        cl3 = 1
        L = LenB(Binary)
        Do While cl1<=L
            pl3 = pl3 & Chr(AscB(MidB(Binary,cl1,1)))
            cl1 = cl1 + 1
            cl3 = cl3 + 1
            if cl3>300 then
                pl2 = pl2 & pl3
                pl3 = ""
                cl3 = 1
                cl2 = cl2 + 1
                if cl2>200 then
                    pl1 = pl1 & pl2
                    pl2 = ""
                    cl2 = 1
                End If
            End If
        Loop
        BinaryToString = pl1 & pl2 & pl3
    End Function

    Private Function StringToBinary(String)
        Dim I, B
        For I=1 to len(String)
            B = B & ChrB(Asc(Mid(String,I,1)))
        Next
        StringToBinary = B
    End Function

End Class

Class FileData
    Private m_fileName
    Private m_contentType
    Private m_BinaryContents
    Private m_AsciiContents
    Private m_imageWidth
    Private m_imageHeight
    Private m_checkImage
    Private m_formFieldName
    Private m_contentDisposition

    Public Property Get FormFieldName
        FormFieldName = m_formFieldName
    End Property

    Public Property Let FormFieldName(sFieldName)
        m_formFieldName = sFieldName
    End Property

    Public Property Get ContentDisposition
        ContentDisposition = m_contentDisposition
    End Property

    Public Property Let ContentDisposition(sContentDisposition)
        m_contentDisposition = sContentDisposition
    End Property

    Public Property Get FileName
        FileName = m_fileName
    End Property

    Public Property Get ContentType
        ContentType = m_contentType
    End Property

    Public Property Get ImageWidth
        If m_checkImage=False Then Call CheckImageDimensions
        ImageWidth = m_imageWidth
    End Property

    Public Property Get ImageHeight
        If m_checkImage=False Then Call CheckImageDimensions
        ImageHeight = m_imageHeight
    End Property

    Public Property Let FileName(ByVal strName)
        strName = Replace(strName, "/", "\")
        Dim arrTemp : arrTemp = Split(strName, "\")
        m_fileName = arrTemp(UBound(arrTemp))
    End Property

    Public Property Let CheckImage(blnCheck)
        m_checkImage = blnCheck
    End Property

    Public Property Let ContentType(strType)
        m_contentType = strType
    End Property

    Public Property Let Contents(strData)
        m_BinaryContents = strData
        m_AsciiContents = RSBinaryToString(m_BinaryContents)
    End Property

    Public Property Get Size
        Size = LenB(m_BinaryContents)
    End Property

    Private Sub CheckImageDimensions
        Dim width, height, colors
        Dim strType

        '''If gfxSpex(BinaryToAscii(m_BinaryContents), width, height, colors, strType) = true then
        If gfxSpex(m_AsciiContents, width, height, colors, strType) = true then
            m_imageWidth = width
            m_imageHeight = height
        End If
        m_checkImage = True
    End Sub

    Private Sub Class_Initialize
        m_imageWidth = -1
        m_imageHeight = -1
        m_checkImage = False
    End Sub

    Public Sub SaveToDisk(strFolderPath, ByRef strNewFileName)
        Dim strPath, objFSO, objFile
        Dim i, time1, time2
        Dim objStream, strExtension

        strPath = strFolderPath&"\"
        If Len(strNewFileName)=0 Then
            strPath = strPath & m_fileName
        Else  
            strExtension = GetExtension(strNewFileName)
            If Len(strExtension)=0 Then
                strNewFileName = strNewFileName & "." & GetExtension(m_fileName)
            End If
            strPath = strPath & strNewFileName
        End If

        time1 = CDbl(Timer)

        Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
        Set objFile = objFSO.CreateTextFile(strPath)

        objFile.Write(m_AsciiContents)

        '''For i=1 to LenB(m_BinaryContents)
        '''    objFile.Write chr(AscB(MidB(m_BinaryContents, i, 1)))
        '''Next          

        time2 = CDbl(Timer)

        objFile.Close
        Set objFile=Nothing
        Set objFSO=Nothing
    End Sub

    Private Function GetExtension(strPath)
        Dim arrTemp
        arrTemp = Split(strPath, ".")
        GetExtension = ""
        If UBound(arrTemp)>0 Then
            GetExtension = arrTemp(UBound(arrTemp))
        End If
    End Function

    Private Function RSBinaryToString(xBinary)
        'Antonin Foller, http://www.motobit.com
        'RSBinaryToString converts binary data (VT_UI1 | VT_ARRAY Or MultiByte string)
        'to a string (BSTR) using ADO recordset

        Dim Binary
        'MultiByte data must be converted To VT_UI1 | VT_ARRAY first.
        If vartype(xBinary)=8 Then Binary = MultiByteToBinary(xBinary) Else Binary = xBinary

        Dim RS, LBinary
        Const adLongVarChar = 201
        Set RS = CreateObject("ADODB.Recordset")
        LBinary = LenB(Binary)

        If LBinary>0 Then
            RS.Fields.Append "mBinary", adLongVarChar, LBinary
            RS.Open
            RS.AddNew
            RS("mBinary").AppendChunk Binary 
            RS.Update
            RSBinaryToString = RS("mBinary")
        Else  
            RSBinaryToString = ""
        End If
    End Function

    Function MultiByteToBinary(MultiByte)
        '© 2000 Antonin Foller, http://www.motobit.com
        ' MultiByteToBinary converts multibyte string To real binary data (VT_UI1 | VT_ARRAY)
        ' Using recordset
        Dim RS, LMultiByte, Binary
        Const adLongVarBinary = 205
        Set RS = CreateObject("ADODB.Recordset")
        LMultiByte = LenB(MultiByte)
        If LMultiByte>0 Then
            RS.Fields.Append "mBinary", adLongVarBinary, LMultiByte
            RS.Open
            RS.AddNew
            RS("mBinary").AppendChunk MultiByte & ChrB(0)
            RS.Update
            Binary = RS("mBinary").GetChunk(LMultiByte)
        End If
        MultiByteToBinary = Binary
    End Function

    Private Function BinaryToAscii(strBinary)
        Dim i, result
        result = ""
        For i=1 to LenB(strBinary)
            result = result & chr(AscB(MidB(strBinary, i, 1))) 
        Next
        BinaryToAscii = result
    End Function

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ':::                                                             :::
    ':::  This routine will attempt to identify any filespec passed  :::
    ':::  as a graphic file (regardless of the extension). This will :::
    ':::  work with BMP, GIF, JPG and PNG files.                     :::
    ':::                                                             :::
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ':::          Based on ideas presented by David Crowell          :::
    ':::                   (credit where due)                        :::
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '::: blah blah blah blah blah blah blah blah blah blah blah blah :::
    '::: blah blah blah blah blah blah blah blah blah blah blah blah :::
    '::: blah blah     Copyright *c* MM,  Mike Shaffer     blah blah :::
    '::: bh blah      ALL RIGHTS RESERVED WORLDWIDE      blah blah :::
    '::: blah blah  Permission is granted to use this code blah blah :::
    '::: blah blah   in your projects, as long as this     blah blah :::
    '::: blah blah      copyright notice is included       blah blah :::
    '::: blah blah blah blah blah blah blah blah blah blah blah blah :::
    '::: blah blah blah blah blah blah blah blah blah blah blah blah :::
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ':::                                                             :::
    ':::  This function gets a specified number of bytes from any    :::
    ':::  file, starting at the offset (base 1)                      :::
    ':::                                                             :::
    ':::  Passed:                                                    :::
    ':::       flnm        => Filespec of file to read               :::
    ':::       offset      => Offset at which to start reading       :::
    ':::       bytes       => How many bytes to read                 :::
    ':::                                                             :::
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Private Function GetBytes(flnm, offset, bytes)
        Dim startPos
        If offset=0 Then
            startPos = 1
        Else  
            startPos = offset
        End If
        if bytes = -1 then        ' Get All!
            GetBytes = flnm
        else
            GetBytes = Mid(flnm, startPos, bytes)
        end if
'        Dim objFSO
'        Dim objFTemp
'        Dim objTextStream
'        Dim lngSize
'        
'        Set objFSO = CreateObject("Scripting.FileSystemObject")
'        
'        ' First, we get the filesize
'        Set objFTemp = objFSO.GetFile(flnm)
'        lngSize = objFTemp.Size
'        set objFTemp = nothing
'        
'        fsoForReading = 1
'        Set objTextStream = objFSO.OpenTextFile(flnm, fsoForReading)
'        
'        if offset > 0 then
'            strBuff = objTextStream.Read(offset - 1)
'        end if
'        
'        if bytes = -1 then        ' Get All!
'            GetBytes = objTextStream.Read(lngSize)  'ReadAll
'        else
'            GetBytes = objTextStream.Read(bytes)
'        end if
'        
'        objTextStream.Close
'        set objTextStream = nothing
'        set objFSO = nothing
    End Function

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ':::                                                             :::
    ':::  Functions to convert two bytes to a numeric value (long)   :::
    ':::  (both little-endian and big-endian)                        :::
    ':::                                                             :::
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Private Function lngConvert(strTemp)
        lngConvert = clng(asc(left(strTemp, 1)) + ((asc(right(strTemp, 1)) * 256)))
    end function

    Private Function lngConvert2(strTemp)
        lngConvert2 = clng(asc(right(strTemp, 1)) + ((asc(left(strTemp, 1)) * 256)))
    end function

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ':::                                                             :::
    ':::  This function does most of the real work. It will attempt  :::
    ':::  to read any file, regardless of the extension, and will    :::
    ':::  identify if it is a graphical image.                       :::
    ':::                                                             :::
    ':::  Passed:                                                    :::
    ':::       flnm        => Filespec of file to read               :::
    ':::       width       => width of image                         :::
    ':::       height      => height of image                        :::
    ':::       depth       => color depth (in number of colors)      :::
    ':::       strImageType=> type of image (e.g. GIF, BMP, etc.)    :::
    ':::                                                             :::
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    function gfxSpex(flnm, width, height, depth, strImageType)
        dim strPNG 
        dim strGIF
        dim strBMP
        dim strType
        dim strBuff
        dim lngSize
        dim flgFound
        dim strTarget
        dim lngPos
        dim ExitLoop
        dim lngMarkerSize

        strType = ""
        strImageType = "(unknown)"

        gfxSpex = False

        strPNG = chr(137) & chr(80) & chr(78)
        strGIF = "GIF"
        strBMP = chr(66) & chr(77)

        strType = GetBytes(flnm, 0, 3)

        if strType = strGIF then                ' is GIF
            strImageType = "GIF"
            Width = lngConvert(GetBytes(flnm, 7, 2))
            Height = lngConvert(GetBytes(flnm, 9, 2))
            Depth = 2 ^ ((asc(GetBytes(flnm, 11, 1)) and 7) + 1)
            gfxSpex = True
        elseif left(strType, 2) = strBMP then        ' is BMP
            strImageType = "BMP"
            Width = lngConvert(GetBytes(flnm, 19, 2))
            Height = lngConvert(GetBytes(flnm, 23, 2))
            Depth = 2 ^ (asc(GetBytes(flnm, 29, 1)))
            gfxSpex = True
        elseif strType = strPNG then            ' Is PNG
            strImageType = "PNG"
            Width = lngConvert2(GetBytes(flnm, 19, 2))
            Height = lngConvert2(GetBytes(flnm, 23, 2))
            Depth = getBytes(flnm, 25, 2)
            select case asc(right(Depth,1))
                case 0
                    Depth = 2 ^ (asc(left(Depth, 1)))
                    gfxSpex = True
                case 2
                    Depth = 2 ^ (asc(left(Depth, 1)) * 3)
                    gfxSpex = True
                case 3
                    Depth = 2 ^ (asc(left(Depth, 1)))  '8
                    gfxSpex = True
                case 4
                    Depth = 2 ^ (asc(left(Depth, 1)) * 2)
                    gfxSpex = True
                case 6
                    Depth = 2 ^ (asc(left(Depth, 1)) * 4)
                    gfxSpex = True
                case else
                    Depth = -1
            end select
        else
            strBuff = GetBytes(flnm, 0, -1)        ' Get all bytes from file
            lngSize = len(strBuff)
            flgFound = 0

            strTarget = chr(255) & chr(216) & chr(255)
            flgFound = instr(strBuff, strTarget)

            if flgFound = 0 then
                exit function
            end if

            strImageType = "JPG"
            lngPos = flgFound + 2
            ExitLoop = false

            do while ExitLoop = False and lngPos < lngSize
                do while asc(mid(strBuff, lngPos, 1)) = 255 and lngPos < lngSize
                    lngPos = lngPos + 1
                loop

                if asc(mid(strBuff, lngPos, 1)) < 192 or asc(mid(strBuff, lngPos, 1)) > 195 then
                    lngMarkerSize = lngConvert2(mid(strBuff, lngPos + 1, 2))
                    lngPos = lngPos + lngMarkerSize  + 1
                else
                    ExitLoop = True
                end if
            loop

            if ExitLoop = False then
                Width = -1
                Height = -1
                Depth = -1
            else
                Height = lngConvert2(mid(strBuff, lngPos + 4, 2))
                Width = lngConvert2(mid(strBuff, lngPos + 6, 2))
                Depth = 2 ^ (asc(mid(strBuff, lngPos + 8, 1)) * 8)
                gfxSpex = True
            end if
        end if
    End Function
End Class
%>