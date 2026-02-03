<%@ Language="VBScript" %>
<%
' Test to see what happens in the upload flow
Response.ContentType = "text/plain"

Response.Write "=== Upload Debug Test 2 ===" & vbCrLf

' Check content type
If InStr(Request.ServerVariables("CONTENT_TYPE"), "multipart/form-data") > 0 Then
    Response.Write "Multipart form data detected" & vbCrLf
    
    ' Create stream
    Set StreamRequest = Server.CreateObject("ADODB.Stream")
    StreamRequest.Type = 1 ' adTypeBinary
    StreamRequest.Open
    
    ' Read binary data
    totalBytes = Request.TotalBytes
    Response.Write "TotalBytes: " & totalBytes & vbCrLf
    
    If totalBytes > 0 Then
        binData = Request.BinaryRead(totalBytes)
        StreamRequest.Write binData
        Response.Write "Binary data written to stream" & vbCrLf
        Response.Write "Stream Size: " & StreamRequest.Size & vbCrLf
        
        ' Test dictionary creation and iteration
        Set UploadedFiles = Server.CreateObject("Scripting.Dictionary")
        Response.Write "Dictionary created" & vbCrLf
        
        ' Add simple test data to dictionary
        UploadedFiles.Add "testfile", "test_value"
        Response.Write "Test value added to dictionary" & vbCrLf
        Response.Write "UploadedFiles.Count = " & UploadedFiles.Count & vbCrLf
        
        ' Test iteration over keys
        Response.Write "Testing For Each Keys iteration..." & vbCrLf
        For Each key In UploadedFiles.Keys
            Response.Write "Key: " & key & vbCrLf
        Next
        
        ' Test iteration over items
        Response.Write "Testing For Each Items iteration..." & vbCrLf
        For Each item In UploadedFiles.Items
            Response.Write "Item: " & item & vbCrLf
        Next
        
        ' Test save operation directly with stream
        Response.Write vbCrLf & "Testing direct stream save..." & vbCrLf
        uploadPath = Server.MapPath("/uploads/")
        Response.Write "Upload path: " & uploadPath & vbCrLf
        
        ' Create output stream
        Set streamFile = Server.CreateObject("ADODB.Stream")
        streamFile.Type = 1 'adTypeBinary
        streamFile.Open
        
        ' Position at start of file content (after headers, ~195 bytes)
        startPos = 195
        fileLen = 100
        Response.Write "Setting StreamRequest.Position = " & startPos & vbCrLf
        StreamRequest.Position = startPos
        Response.Write "StreamRequest.Position after set: " & StreamRequest.Position & vbCrLf
        
        Response.Write "Calling CopyTo with length " & fileLen & vbCrLf
        StreamRequest.CopyTo streamFile, fileLen
        
        Response.Write "streamFile.Size after CopyTo: " & streamFile.Size & vbCrLf
        
        filePath = uploadPath & "test_upload.bin"
        Response.Write "Saving to: " & filePath & vbCrLf
        streamFile.SaveToFile filePath, 2 'adSaveCreateOverWrite
        
        streamFile.Close
        Set streamFile = Nothing
        
        Response.Write "File saved successfully!" & vbCrLf
        
        StreamRequest.Close
        Set StreamRequest = Nothing
    End If
Else
    Response.Write "No multipart form data" & vbCrLf
End If

Response.Write vbCrLf & "=== Test Complete ===" & vbCrLf
%>
