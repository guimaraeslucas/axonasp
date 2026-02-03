<%@ Language="VBScript" %>
<%
Response.ContentType = "text/plain"

On Error Resume Next

' Get binary data from request
Dim readBytes
readBytes = 5000

Dim binData
binData = Request.BinaryRead(readBytes)

Response.Write "Bytes Read: " & readBytes & vbCrLf
Response.Write "LenB: " & LenB(binData) & vbCrLf

' Test creating a stream and writing binary data
Dim stream
Set stream = Server.CreateObject("ADODB.Stream")
stream.Type = 1 ' adTypeBinary
stream.Open

' Write the binary data
stream.Write binData

Response.Write "Stream Size: " & stream.Size & vbCrLf
Response.Write "Stream Position: " & stream.Position & vbCrLf

' Save to file
Dim savePath
savePath = Server.MapPath("../uploads/test_upload_debug.bin")
Response.Write "Save Path: " & savePath & vbCrLf

stream.SaveToFile savePath, 2

stream.Close
Set stream = Nothing

If Err.Number <> 0 Then
    Response.Write "Error: " & Err.Description & vbCrLf
End If

Response.Write "Done!" & vbCrLf
%>
