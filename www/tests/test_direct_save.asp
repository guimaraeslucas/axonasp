<%@ Language="VBScript" %>
<%
Response.ContentType = "text/plain"
Response.Write "=== Direct Stream Save Test ===" & vbCrLf & vbCrLf

' Check if multipart
If InStr(Request.ServerVariables("CONTENT_TYPE"), "multipart/form-data") > 0 Then
    Response.Write "Multipart detected" & vbCrLf
    
    ' Read binary data
    totalBytes = Request.TotalBytes
    Response.Write "TotalBytes: " & totalBytes & vbCrLf
    
    If totalBytes > 0 Then
        binData = Request.BinaryRead(totalBytes)
        Response.Write "Binary read OK, length: " & LenB(binData) & vbCrLf
        
        ' Create source stream
        Set srcStream = Server.CreateObject("ADODB.Stream")
        srcStream.Type = 1 ' binary
        srcStream.Open
        srcStream.Write binData
        Response.Write "Source stream created, Size: " & srcStream.Size & vbCrLf
        Response.Flush
        
        ' Find file content boundaries (after headers)
        ' Headers end with \r\n\r\n
        ' For simplicity, start at byte 200 and copy 50 bytes
        startPos = 195
        copyLen = 50
        
        Response.Write "Setting position to " & startPos & vbCrLf
        Response.Flush
        
        srcStream.Position = startPos
        Response.Write "Position set. Current position: " & srcStream.Position & vbCrLf
        Response.Flush
        
        ' Create destination stream
        Set destStream = Server.CreateObject("ADODB.Stream")
        destStream.Type = 1 ' binary
        destStream.Open
        Response.Write "Dest stream created" & vbCrLf
        Response.Flush
        
        ' Copy
        Response.Write "Calling CopyTo with length " & copyLen & vbCrLf
        Response.Flush
        
        srcStream.CopyTo destStream, copyLen
        Response.Write "CopyTo complete. Dest size: " & destStream.Size & vbCrLf
        Response.Flush
        
        ' Save
        savePath = Server.MapPath("/uploads/test_direct_save.bin")
        Response.Write "Saving to: " & savePath & vbCrLf
        Response.Flush
        
        destStream.SaveToFile savePath, 2
        Response.Write "SaveToFile complete!" & vbCrLf
        
        ' Cleanup
        destStream.Close
        srcStream.Close
        Set destStream = Nothing
        Set srcStream = Nothing
        
        Response.Write vbCrLf & "SUCCESS - File saved!" & vbCrLf
    End If
Else
    Response.Write "Not multipart - nothing to test" & vbCrLf
End If

Response.Write vbCrLf & "=== Test Complete ===" & vbCrLf
%>
