<%@ Language="VBScript" %>
<%
Option Explicit

Response.Write "<html><body>"
Response.Write "<h1>Testing G3MD Library</h1>"

Dim md
Set md = Server.CreateObject("G3MD")

If IsObject(md) Then
    Response.Write "<p>G3MD object created successfully.</p>"
Else
    Response.Write "<p style='color:red'>Failed to create G3MD object.</p>"
End If

Dim source, result
source = "# Test Heading" & vbCrLf & "This is a **bold** text and an *italic* text." & vbCrLf & "- Item 1" & vbCrLf & "- Item 2"

Response.Write "<h2>Default Processing</h2>"
result = md.Process(source)
Response.Write "<div style='border:1px solid #ccc; padding: 10px;'>" & result & "</div>"

Response.Write "<h2>Testing GFM (Tables)</h2>"
source = "| Column 1 | Column 2 |" & vbCrLf & "|---|---|" & vbCrLf & "| Value 1 | Value 2 |"
result = md.Process(source)
Response.Write "<div style='border:1px solid #ccc; padding: 10px;'>" & result & "</div>"

Response.Write "<h2>Testing HardWraps</h2>"
source = "Line 1" & vbCrLf & "Line 2"
Response.Write "<h3>HardWraps = False (Default)</h3>"
md.HardWraps = False
result = md.Process(source)
Response.Write "<div style='border:1px solid #ccc; padding: 10px;'>" & result & "</div>"

Response.Write "<h3>HardWraps = True</h3>"
md.HardWraps = True
result = md.Process(source)
Response.Write "<div style='border:1px solid #ccc; padding: 10px;'>" & result & "</div>"

Response.Write "<h2>Testing Unsafe</h2>"
source = "Some text <script>alert('test');</script>"
Response.Write "<h3>Unsafe = False (Default)</h3>"
md.Unsafe = False
result = md.Process(source)
' Escape it so we can see what was rendered
Response.Write "<div style='border:1px solid #ccc; padding: 10px;'>" & Server.HTMLEncode(result) & "</div>"

Response.Write "<h3>Unsafe = True</h3>"
md.Unsafe = True
result = md.Process(source)
' Escape it so we can see what was rendered
Response.Write "<div style='border:1px solid #ccc; padding: 10px;'>" & Server.HTMLEncode(result) & "</div>"

Response.Write "</body></html>"
%>