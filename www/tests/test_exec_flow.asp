<%@ Language="VBScript" %>
<%
' Test that exactly reproduces aspL.exec flow
Response.ContentType = "text/html"
Response.Buffer = True
%>
<!DOCTYPE html>
<html>
<head><title>Exec Flow Test</title></head>
<body>
<h1>Testing aspL.exec Flow</h1>
<%

On Error Resume Next

Dim uploaderPath
uploaderPath = Server.MapPath("../asplite/plugins/uploader/uploader.asp")
Response.Write "<p>Path: " & uploaderPath & "</p>" & vbCrLf
Response.Flush

' Step 1: stream(path, false, "")
Response.Write "<h2>Step 1: stream() function...</h2>" & vbCrLf
Response.Flush

Dim objStream, content
Set objStream = Server.CreateObject("ADODB.Stream")
objStream.CharSet = "utf-8"
objStream.Open
objStream.Type = 2 'adTypeText
objStream.LoadFromFile(uploaderPath)
content = objStream.ReadText()
Set objStream = Nothing

Response.Write "<p>Content loaded: " & Len(content) & " chars</p>" & vbCrLf

If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ERROR stream: " & Err.Description & "</p>" & vbCrLf
    Err.Clear
    Response.End
End If
Response.Flush

' Step 2: removeCRB()
Response.Write "<h2>Step 2: removeCRB() function...</h2>" & vbCrLf
Response.Flush

Dim code
code = content
code = Replace(code, "<" & "%", "", 1, -1, 1)
code = Replace(code, "%" & ">", "", 1, -1, 1)

Response.Write "<p>Code after removeCRB: " & Len(code) & " chars</p>" & vbCrLf
Response.Write "<p>First 200: <pre>" & Server.HTMLEncode(Left(code, 200)) & "</pre></p>" & vbCrLf
Response.Write "<p>Last 200: <pre>" & Server.HTMLEncode(Right(code, 200)) & "</pre></p>" & vbCrLf

If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ERROR removeCRB: " & Err.Description & "</p>" & vbCrLf
    Err.Clear
    Response.End
End If
Response.Flush

' Step 3: ExecuteGlobal
Response.Write "<h2>Step 3: ExecuteGlobal...</h2>" & vbCrLf
Response.Write "<p>Time: " & Now() & "</p>" & vbCrLf
Response.Flush

ExecuteGlobal code

Response.Write "<p>Completed at: " & Now() & "</p>" & vbCrLf

If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ERROR ExecuteGlobal: " & Err.Number & " - " & Err.Description & "</p>" & vbCrLf
    Err.Clear
Else
    Response.Write "<p style='color:green'>SUCCESS!</p>" & vbCrLf
End If

On Error Goto 0

Response.Write "<h2>Test complete</h2>" & vbCrLf

%>
</body>
</html>
