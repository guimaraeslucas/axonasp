<%@ Language="VBScript" %>
<%
Response.ContentType = "text/html"
Response.Write "<h2>Full Upload Simulation Test</h2>"

' Check if we have a POST
If Request.TotalBytes < 1 Then
    Response.Write "<p>No data received. POST something to this endpoint.</p>"
    Response.End
End If

Response.Write "<p>TotalBytes: " & Request.TotalBytes & "</p>"
Response.Flush

' Exactly as uploader.asp does
Dim VarArrayBinRequest, StreamRequest, internalChunkSize, readBytes, tmpBinRequest

internalChunkSize = 200000

' Create the stream
Response.Write "<p>Creating ADODB.Stream...</p>"
Response.Flush
Set StreamRequest = Server.CreateObject("ADODB.Stream")
StreamRequest.Type = 1 ' adTypeBinary
StreamRequest.Open
Response.Write "<p>Stream created and opened.</p>"
Response.Flush

' Copy binary request
Response.Write "<p>Starting BinaryRead loop...</p>"
Response.Flush
readBytes = internalChunkSize
VarArrayBinRequest = Request.BinaryRead(readBytes)
Response.Write "<p>First read: readBytes=" & readBytes & ", got " & LenB(VarArrayBinRequest) & " bytes</p>"
Response.Flush

VarArrayBinRequest = MidB(VarArrayBinRequest, 1, LenB(VarArrayBinRequest))
Response.Write "<p>After MidB: " & LenB(VarArrayBinRequest) & " bytes</p>"
Response.Flush

Dim loopCount : loopCount = 0
Do Until readBytes < 1
    loopCount = loopCount + 1
    If loopCount > 10 Then
        Response.Write "<p><b>SAFETY EXIT at 10 loops</b></p>"
        Exit Do
    End If
    tmpBinRequest = Request.BinaryRead(readBytes)
    Response.Write "<p>Loop " & loopCount & ": readBytes=" & readBytes & ", got " & LenB(tmpBinRequest) & " bytes</p>"
    Response.Flush
    If readBytes > 0 Then
        VarArrayBinRequest = VarArrayBinRequest & MidB(tmpBinRequest, 1, LenB(tmpBinRequest))
    End If
Loop
Response.Write "<p>Loop finished. Total data: " & LenB(VarArrayBinRequest) & " bytes</p>"
Response.Flush

' Write to stream
Response.Write "<p>Writing to stream...</p>"
Response.Flush
StreamRequest.Write VarArrayBinRequest
Response.Write "<p>Stream write complete. Stream size: " & StreamRequest.Size & "</p>"
Response.Flush

' Now test token finding (exactly as uploader does)
Response.Write "<hr><p><b>Testing Token Finding</b></p>"
Response.Flush

' String2Byte function
Function String2Byte(sString)
    Dim i
    For i = 1 to Len(sString)
       String2Byte = String2Byte & ChrB(AscB(Mid(sString,i,1)))
    Next
End Function

Dim tNewLine, tTerm, vDataSep, nCurPos
tNewLine = String2Byte(Chr(13))
tTerm = String2Byte("--")

Response.Write "<p>tNewLine: " & LenB(tNewLine) & " bytes, tTerm: " & LenB(tTerm) & " bytes</p>"
Response.Flush

' Find first newline - InStrB
Response.Write "<p>Looking for newline in data...</p>"
Response.Flush
nCurPos = InStrB(1, VarArrayBinRequest, tNewLine, 0)
Response.Write "<p>First newline at position: " & nCurPos & "</p>"
Response.Flush

If nCurPos > 0 Then
    vDataSep = MidB(VarArrayBinRequest, 1, nCurPos-1)
    Response.Write "<p>Data separator: " & LenB(vDataSep) & " bytes</p>"
    Response.Flush
    
    ' Try to convert to string for display
    Dim sepStr, j
    sepStr = ""
    For j = 1 To LenB(vDataSep)
        sepStr = sepStr & Chr(AscB(MidB(vDataSep, j, 1)))
    Next
    Response.Write "<p>Separator string: [" & Server.HTMLEncode(sepStr) & "]</p>"
    Response.Flush
End If

' Find terminator
Dim nLastSepPos
nLastSepPos = InStrB(1, VarArrayBinRequest, vDataSep & tTerm, 0)
Response.Write "<p>Last separator position (with --): " & nLastSepPos & "</p>"
Response.Flush

Response.Write "<hr><p><b>Test Complete!</b></p>"

' Cleanup
StreamRequest.Close
Set StreamRequest = Nothing
%>
