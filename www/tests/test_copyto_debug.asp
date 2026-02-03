<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.ContentType = "text/plain"
Response.Write "=== COPYTO DEBUG ===" & vbCrLf

On Error Resume Next

' Create source stream with some data
Dim srcStream
Set srcStream = Server.CreateObject("ADODB.Stream")
srcStream.Type = 2  ' text
srcStream.Open
srcStream.WriteText "Hello World Test Data for CopyTo"
srcStream.Flush
Response.Write "Source Size: " & srcStream.Size & vbCrLf

' Create destination stream
Dim dstStream
Set dstStream = Server.CreateObject("ADODB.Stream")
dstStream.Charset = "utf-8"
dstStream.Mode = 3
dstStream.Type = 1  ' binary
dstStream.Open

Response.Write "Dest initial Position: " & dstStream.Position & vbCrLf
Response.Write "Dest initial Size: " & dstStream.Size & vbCrLf

' Test CopyTo from position 7 (skipping "Hello W")
srcStream.Position = 7
Response.Write "Source Position: " & srcStream.Position & vbCrLf
srcStream.CopyTo dstStream, 5  ' copy "orld "
dstStream.Flush

Response.Write "Dest Size after CopyTo: " & dstStream.Size & vbCrLf
Response.Write "Dest Position after CopyTo: " & dstStream.Position & vbCrLf

' Set position to 0 for reading
dstStream.Position = 0
Response.Write "Dest Position after reset: " & dstStream.Position & vbCrLf

' Change to text type
dstStream.Type = 2
Response.Write "Dest Type after change: " & dstStream.Type & vbCrLf
Response.Write "Dest Size after type change: " & dstStream.Size & vbCrLf
Response.Write "Dest Position after type change: " & dstStream.Position & vbCrLf

' Read the destination - use WITHOUT parentheses to test property access
Response.Write "About to call ReadText (no parens)..." & vbCrLf
Dim result
result = dstStream.ReadText
Response.Write "ReadText called" & vbCrLf
Response.Write "Read result: [" & result & "]" & vbCrLf
Response.Write "Result length: " & Len(result) & vbCrLf

If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf
    Err.Clear
End If

srcStream.Close
dstStream.Close
Set srcStream = Nothing
Set dstStream = Nothing

Response.Write vbCrLf & "=== DONE ===" & vbCrLf
%>
