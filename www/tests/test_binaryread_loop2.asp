<%@ Language="VBScript" %>
<%
Response.ContentType = "text/html"
Response.Buffer = False

Response.Write "<h2>BinaryRead Loop Test - Exactly as Uploader</h2>"
Response.Flush

If Request.TotalBytes < 1 Then
    Response.Write "<p>No data.</p>"
    Response.End
End If

Response.Write "<p>TotalBytes: " & Request.TotalBytes & "</p>"
Response.Flush

Dim readBytes, VarArrayBinRequest, tmpBinRequest
Dim internalChunkSize
internalChunkSize = 200000

' Exactly as uploader does - no safety limit
Response.Write "<p>Starting loop exactly as uploader...</p>"
Response.Flush

readBytes = internalChunkSize
VarArrayBinRequest = Request.BinaryRead(readBytes)
Response.Write "<p>First read: readBytes=" & readBytes & ", got " & LenB(VarArrayBinRequest) & "</p>"
Response.Flush

VarArrayBinRequest = MidB(VarArrayBinRequest, 1, LenB(VarArrayBinRequest))

Dim loopCount : loopCount = 0
Response.Write "<p>Entering loop...</p>" : Response.Flush

Do Until readBytes < 1
    loopCount = loopCount + 1
    Response.Write "<p>Loop " & loopCount & ": readBytes=" & readBytes & "</p>" : Response.Flush
    
    ' SAFETY: Exit if too many loops
    If loopCount > 100 Then
        Response.Write "<p style='color:red'>INFINITE LOOP DETECTED! Exiting.</p>"
        Exit Do
    End If
    
    tmpBinRequest = Request.BinaryRead(readBytes)
    Response.Write "<p>  After BinaryRead: readBytes=" & readBytes & ", got " & LenB(tmpBinRequest) & "</p>" : Response.Flush
    
    If readBytes > 0 Then
        VarArrayBinRequest = VarArrayBinRequest & MidB(tmpBinRequest, 1, LenB(tmpBinRequest))
    End If
Loop

Response.Write "<p>Loop finished after " & loopCount & " iterations</p>"
Response.Write "<p>Total data: " & LenB(VarArrayBinRequest) & " bytes</p>"
Response.Write "<p>Test complete!</p>"
%>
