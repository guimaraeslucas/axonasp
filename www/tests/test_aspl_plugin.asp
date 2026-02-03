<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include virtual="/asplite/asplite.asp"-->
<%
Response.ContentType = "text/html"
Response.Buffer = False

Response.Write "<h2>aspL.plugin Test</h2>"
Response.Flush

Response.Write "<p>About to load uploader plugin...</p>"
Response.Write "<p>Time: " & Now() & "</p>"
Response.Flush

On Error Resume Next

' This is what we're testing
Dim uploader
Set uploader = aspL.plugin("uploader")

Response.Write "<p>Time: " & Now() & "</p>"
Response.Flush

If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Error: " & Err.Number & " - " & Err.Description & "</p>"
    Err.Clear
Else
    Response.Write "<p style='color:green'>plugin() completed!</p>"
    If Not uploader Is Nothing Then
        Response.Write "<p>uploader object type: " & TypeName(uploader) & "</p>"
    Else
        Response.Write "<p>uploader is Nothing!</p>"
    End If
End If

On Error GoTo 0

Response.Write "<hr><p><b>Test complete!</b></p>"
Set aspL = Nothing
%>
