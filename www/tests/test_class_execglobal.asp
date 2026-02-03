<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include virtual="/asplite/asplite.asp"-->
<%
Response.ContentType = "text/html"
Response.Buffer = False

Response.Write "<h2>Simple Class via ExecuteGlobal Test</h2>"
Response.Flush

' Define a simple class with BinaryRead
Dim testCode
testCode = "Class TestUploader" & vbCrLf
testCode = testCode & "  Private m_data" & vbCrLf
testCode = testCode & "  Public Sub Read()" & vbCrLf
testCode = testCode & "    Response.Write ""<p>Inside Read()</p>"" : Response.Flush" & vbCrLf
testCode = testCode & "    Dim rb : rb = 200000" & vbCrLf
testCode = testCode & "    Response.Write ""<p>Before BinaryRead: rb="" & rb & ""</p>"" : Response.Flush" & vbCrLf
testCode = testCode & "    m_data = Request.BinaryRead(rb)" & vbCrLf
testCode = testCode & "    Response.Write ""<p>After BinaryRead: rb="" & rb & "", got="" & LenB(m_data) & ""</p>"" : Response.Flush" & vbCrLf
testCode = testCode & "    Dim loop1 : loop1 = 0" & vbCrLf
testCode = testCode & "    Do Until rb < 1" & vbCrLf
testCode = testCode & "      loop1 = loop1 + 1" & vbCrLf
testCode = testCode & "      Response.Write ""<p>Loop "" & loop1 & "": rb="" & rb & ""</p>"" : Response.Flush" & vbCrLf
testCode = testCode & "      If loop1 > 5 Then Exit Do" & vbCrLf
testCode = testCode & "      Dim tmp" & vbCrLf
testCode = testCode & "      tmp = Request.BinaryRead(rb)" & vbCrLf
testCode = testCode & "      Response.Write ""<p>After loop read: rb="" & rb & ""</p>"" : Response.Flush" & vbCrLf
testCode = testCode & "    Loop" & vbCrLf
testCode = testCode & "    Response.Write ""<p>Loop complete, total="" & LenB(m_data) & ""</p>"" : Response.Flush" & vbCrLf
testCode = testCode & "  End Sub" & vbCrLf
testCode = testCode & "End Class" & vbCrLf

Response.Write "<pre>" & Server.HTMLEncode(testCode) & "</pre>"
Response.Flush

If Request.TotalBytes < 1 Then
    Response.Write "<p>No POST data.</p>"
    Response.End
End If

Response.Write "<p>TotalBytes: " & Request.TotalBytes & "</p>"
Response.Flush

' Execute the class definition
Response.Write "<hr><p>Executing class definition...</p>"
Response.Flush

ExecuteGlobal testCode

Response.Write "<p style='color:green'>Class defined!</p>"
Response.Flush

' Create instance
Response.Write "<hr><p>Creating instance...</p>"
Response.Flush

Dim obj
Set obj = New TestUploader
Response.Write "<p style='color:green'>Instance created: " & TypeName(obj) & "</p>"
Response.Flush

' Call Read
Response.Write "<hr><p>Calling obj.Read()...</p>"
Response.Flush

obj.Read

Response.Write "<hr><p><b>Test complete!</b></p>"
Set obj = Nothing
%>
