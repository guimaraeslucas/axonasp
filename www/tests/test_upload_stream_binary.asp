<%@ Language="VBScript" %>
<%
Response.ContentType = "text/plain"
Response.Write "=== UPLOAD STREAM BINARY TEST ===" & vbCrLf & vbCrLf

If Request.TotalBytes < 1 Then
    Response.Write "No data. POST multipart data to this endpoint." & vbCrLf
    Response.End
End If

' Read all binary data
Dim binData, readBytes
readBytes = Request.TotalBytes
binData = Request.BinaryRead(readBytes)

Response.Write "Read " & LenB(binData) & " bytes" & vbCrLf & vbCrLf

' Create a stream and write binary data
Dim StreamRequest
Set StreamRequest = Server.CreateObject("ADODB.Stream")
StreamRequest.Type = 2 ' adTypeText
StreamRequest.Open
StreamRequest.WriteText binData
StreamRequest.Flush

Response.Write "Stream Size: " & StreamRequest.Size & vbCrLf
Response.Write "Stream Position: " & StreamRequest.Position & vbCrLf & vbCrLf

' Find Content-Disposition in the binary data
Dim tContentDisp, pos
tContentDisp = ChrB(67) & ChrB(111) & ChrB(110) & ChrB(116) & ChrB(101) & ChrB(110) & ChrB(116) & ChrB(45) & ChrB(68) & ChrB(105) & ChrB(115) & ChrB(112) & ChrB(111) & ChrB(115) & ChrB(105) & ChrB(116) & ChrB(105) & ChrB(111) & ChrB(110)
' That's "Content-Disposition" as bytes

pos = InstrB(1, binData, tContentDisp)
Response.Write "Content-Disposition found at position: " & pos & vbCrLf

' Test reading from stream at that position
If pos > 0 Then
    ' Create temp stream for conversion
    Dim objStream
    Set objStream = Server.CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Mode = 3
    objStream.Type = 1
    objStream.Open
    
    ' Set source position (convert from 1-based to 0-based)
    StreamRequest.Position = pos - 1
    Response.Write "Source Position set to: " & StreamRequest.Position & vbCrLf
    
    ' Copy 19 characters ("Content-Disposition")
    StreamRequest.CopyTo objStream, 19
    objStream.Flush
    
    Response.Write "Dest Size: " & objStream.Size & vbCrLf
    
    ' Read the result
    objStream.Position = 0
    objStream.Type = 2
    
    Dim result
    result = objStream.ReadText
    Response.Write "Read result: [" & result & "]" & vbCrLf
    
    objStream.Close
    Set objStream = Nothing
End If

StreamRequest.Close
Set StreamRequest = Nothing

Response.Write vbCrLf & "=== DONE ===" & vbCrLf
%>
