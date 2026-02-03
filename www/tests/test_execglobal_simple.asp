<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.ContentType = "text/html"
Response.Buffer = False

Response.Write "<h2>ExecuteGlobal Timing Test (NO INCLUDE)</h2>"
Response.Flush

' Load original uploader
Dim objStream, fileContent
Set objStream = Server.CreateObject("ADODB.Stream")
objStream.CharSet = "utf-8"
objStream.Open
objStream.Type = 2
objStream.LoadFromFile Server.MapPath("/asplite/plugins/uploader/uploader.asp")
fileContent = objStream.ReadText()
objStream.Close
Set objStream = Nothing

Response.Write "<p>Original file: " & Len(fileContent) & " chars</p>"
Response.Flush

' Remove code blocks
fileContent = Replace(fileContent, "<" & "%", "", 1, -1, 1)
fileContent = Replace(fileContent, "%" & ">", "", 1, -1, 1)

Response.Write "<p>Clean file: " & Len(fileContent) & " chars</p>"
Response.Flush

Response.Write "<hr><p>About to call ExecuteGlobal...</p>"
Response.Flush

' Add some debug before ExecuteGlobal
Response.Write "<p>Time: " & Now() & "</p>"
Response.Flush

On Error Resume Next
ExecuteGlobal fileContent

Response.Write "<p>Time: " & Now() & "</p>"
Response.Flush

If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ExecuteGlobal error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
Else
    Response.Write "<p style='color:green'>ExecuteGlobal completed!</p>"
End If
Response.Flush

On Error GoTo 0

Response.Write "<hr><p><b>ExecuteGlobal test complete!</b></p>"
%>
