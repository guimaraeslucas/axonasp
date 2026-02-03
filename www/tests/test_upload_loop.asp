<%@ Language="VBScript" %>
<%
Response.ContentType = "text/html"
Response.Write "<h2>Upload Loop Test - Simulating Uploader Plugin</h2>"

Dim internalChunkSize, readBytes, VarArrayBinRequest, tmpBinRequest
Dim loopCount, totalBytes

internalChunkSize = 200000
loopCount = 0
totalBytes = 0

Response.Write "<p>Total content length: " & Request.TotalBytes & "</p>"
Response.Write "<p>internalChunkSize: " & internalChunkSize & "</p>"
Response.Write "<hr>"

' First read - exactly as uploader does
readBytes = internalChunkSize
Response.Write "<p>Before first BinaryRead: readBytes = " & readBytes & "</p>"
VarArrayBinRequest = Request.BinaryRead(readBytes)
Response.Write "<p>After first BinaryRead: readBytes = " & readBytes & ", got " & LenB(VarArrayBinRequest) & " bytes</p>"
Response.Flush

' Process with MidB as uploader does
VarArrayBinRequest = MidB(VarArrayBinRequest, 1, LenB(VarArrayBinRequest))
Response.Write "<p>After MidB: length = " & LenB(VarArrayBinRequest) & "</p>"
Response.Flush

' Loop exactly as uploader
Response.Write "<p>Starting loop...</p>"
Response.Flush

Do Until readBytes < 1
    loopCount = loopCount + 1
    Response.Write "<p>Loop " & loopCount & " - readBytes before = " & readBytes & "</p>"
    Response.Flush
    
    If loopCount > 10 Then
        Response.Write "<p><strong>LOOP LIMIT REACHED - Breaking to avoid timeout</strong></p>"
        Exit Do
    End If
    
    tmpBinRequest = Request.BinaryRead(readBytes)
    Response.Write "<p>Loop " & loopCount & " - After BinaryRead: readBytes = " & readBytes & ", got " & LenB(tmpBinRequest) & " bytes</p>"
    Response.Flush
    
    If readBytes > 0 Then
        VarArrayBinRequest = VarArrayBinRequest & MidB(tmpBinRequest, 1, LenB(tmpBinRequest))
        totalBytes = totalBytes + LenB(tmpBinRequest)
        Response.Write "<p>Loop " & loopCount & " - Appended, total now: " & LenB(VarArrayBinRequest) & "</p>"
        Response.Flush
    End If
Loop

Response.Write "<hr>"
Response.Write "<p><strong>Results:</strong></p>"
Response.Write "<p>Loop iterations: " & loopCount & "</p>"
Response.Write "<p>Total bytes read: " & LenB(VarArrayBinRequest) & "</p>"
Response.Write "<p>Test complete!</p>"
%>
