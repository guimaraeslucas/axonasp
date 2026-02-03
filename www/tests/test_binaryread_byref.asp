<%@ Language="VBScript" %>
<%
Response.ContentType = "text/plain"
Response.Write "=== BinaryRead ByRef Test ===" & vbCrLf & vbCrLf

totalBytes = Request.TotalBytes
Response.Write "TotalBytes: " & totalBytes & vbCrLf

If totalBytes > 0 Then
    ' Test 1: First read with chunk size
    Response.Write vbCrLf & "Test 1: Reading with chunk size..." & vbCrLf
    
    chunkSize = 100
    Response.Write "Initial chunkSize: " & chunkSize & vbCrLf
    
    data1 = Request.BinaryRead(chunkSize)
    
    Response.Write "After read 1: chunkSize = " & chunkSize & vbCrLf
    Response.Write "Data1 length: " & LenB(data1) & vbCrLf
    
    ' Test 2: Second read
    Response.Write vbCrLf & "Test 2: Second read..." & vbCrLf
    Response.Write "Before read 2: chunkSize = " & chunkSize & vbCrLf
    
    data2 = Request.BinaryRead(chunkSize)
    
    Response.Write "After read 2: chunkSize = " & chunkSize & vbCrLf
    Response.Write "Data2 length: " & LenB(data2) & vbCrLf
    
    ' Test 3: Third read
    Response.Write vbCrLf & "Test 3: Third read..." & vbCrLf
    Response.Write "Before read 3: chunkSize = " & chunkSize & vbCrLf
    
    data3 = Request.BinaryRead(chunkSize)
    
    Response.Write "After read 3: chunkSize = " & chunkSize & vbCrLf
    Response.Write "Data3 length: " & LenB(data3) & vbCrLf
    
    ' Test 4: Loop pattern (what uploader does)
    Response.Write vbCrLf & "Test 4: Loop pattern simulation..." & vbCrLf
    loopCount = 0
    readBytes = 100
    maxLoops = 10
    
    Do Until readBytes < 1 Or loopCount >= maxLoops
        loopCount = loopCount + 1
        Response.Write "Loop " & loopCount & ": readBytes before = " & readBytes
        
        tmpData = Request.BinaryRead(readBytes)
        
        Response.Write ", after = " & readBytes & ", got " & LenB(tmpData) & " bytes" & vbCrLf
        Response.Flush
    Loop
    
    Response.Write "Loop exited after " & loopCount & " iterations" & vbCrLf
    
Else
    Response.Write "No data to read (TotalBytes = 0)" & vbCrLf
End If

Response.Write vbCrLf & "=== Test Complete ===" & vbCrLf
%>
