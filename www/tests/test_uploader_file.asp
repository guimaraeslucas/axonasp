<%@ Language="VBScript" %>
<%
' Test reading and parsing uploader.asp file
Response.ContentType = "text/html"
Response.Buffer = True
%>
<!DOCTYPE html>
<html>
<head><title>Uploader File Test</title></head>
<body>
<h1>Testing Uploader.asp File Load</h1>
<%

On Error Resume Next

' Step 1: Read the file directly using FSO
Response.Write "<h2>Step 1: Reading file with FSO...</h2>" & vbCrLf
Response.Flush

Dim fso, uploaderPath, uploaderFile, uploaderContent
Set fso = Server.CreateObject("Scripting.FileSystemObject")
uploaderPath = Server.MapPath("../asplite/plugins/uploader/uploader.asp")
Response.Write "<p>Path: " & uploaderPath & "</p>" & vbCrLf

If fso.FileExists(uploaderPath) Then
    Response.Write "<p style='color:green'>File exists</p>" & vbCrLf
    Set uploaderFile = fso.OpenTextFile(uploaderPath, 1)
    uploaderContent = uploaderFile.ReadAll()
    uploaderFile.Close
    Set uploaderFile = Nothing
    Response.Write "<p>Content length: " & Len(uploaderContent) & " chars</p>" & vbCrLf
Else
    Response.Write "<p style='color:red'>File NOT found!</p>" & vbCrLf
    Response.End
End If
Set fso = Nothing

If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ERROR reading file: " & Err.Description & "</p>" & vbCrLf
    Err.Clear
End If

Response.Flush

' Step 2: Remove ASP tags
Response.Write "<h2>Step 2: Removing ASP tags...</h2>" & vbCrLf
Response.Flush

Dim code
' Remove <% at beginning and %> at end
code = uploaderContent
If Left(code, 2) = "<%" Then code = Mid(code, 3)
' Find and remove %> at end (handling possible line breaks)
Dim endPos
endPos = InStrRev(code, "%>")
If endPos > 0 Then code = Left(code, endPos - 1)

' Trim whitespace
code = Trim(code)

Response.Write "<p>Code length after tag removal: " & Len(code) & " chars</p>" & vbCrLf
Response.Write "<p>First 200 chars: <pre>" & Server.HTMLEncode(Left(code, 200)) & "</pre></p>" & vbCrLf
Response.Write "<p>Last 200 chars: <pre>" & Server.HTMLEncode(Right(code, 200)) & "</pre></p>" & vbCrLf
Response.Flush

' Step 3: Try to parse it
Response.Write "<h2>Step 3: ExecuteGlobal on the code...</h2>" & vbCrLf
Response.Write "<p>Starting at: " & Now() & "</p>" & vbCrLf
Response.Flush

ExecuteGlobal code

If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ERROR executing: " & Err.Number & " - " & Err.Description & "</p>" & vbCrLf
    Err.Clear
Else
    Response.Write "<p style='color:green'>SUCCESS! Code executed without error</p>" & vbCrLf
    Response.Write "<p>Finished at: " & Now() & "</p>" & vbCrLf
End If

On Error Goto 0

Response.Write "<h2>Test complete</h2>" & vbCrLf

%>
</body>
</html>
