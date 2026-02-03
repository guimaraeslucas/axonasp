<%@ Language="VBScript" %>
<%
' Test BinaryRead functionality
Response.ContentType = "text/plain"

Dim totalBytes, readBytes, data, loopCount

totalBytes = Request.TotalBytes
Response.Write "TotalBytes: " & totalBytes & vbCrLf

If totalBytes > 0 Then
    readBytes = totalBytes
    loopCount = 0
    
    Response.Write "Attempting BinaryRead..." & vbCrLf
    
    ' Try to read the binary data
    On Error Resume Next
    data = Request.BinaryRead(readBytes)
    If Err.Number <> 0 Then
        Response.Write "Error in BinaryRead: " & Err.Description & vbCrLf
        Err.Clear
    Else
        Response.Write "BinaryRead returned: " & TypeName(data) & vbCrLf
        Response.Write "Data length: " & LenB(data) & vbCrLf
    End If
    On Error Goto 0
    
    ' Test the loop pattern used by uploader
    Response.Write vbCrLf & "Testing loop pattern..." & vbCrLf
    
    Do Until readBytes < 1
        loopCount = loopCount + 1
        If loopCount > 5 Then
            Response.Write "Breaking after 5 iterations to prevent infinite loop" & vbCrLf
            Exit Do
        End If
        Response.Write "Loop iteration " & loopCount & ", readBytes=" & readBytes & vbCrLf
        data = Request.BinaryRead(readBytes)
        Response.Write "After BinaryRead, readBytes=" & readBytes & ", data type=" & TypeName(data) & vbCrLf
    Loop
    
    Response.Write "Loop completed after " & loopCount & " iterations" & vbCrLf
Else
    Response.Write "No data in request body" & vbCrLf
End If

Response.Write vbCrLf & "Test complete." & vbCrLf
%>
