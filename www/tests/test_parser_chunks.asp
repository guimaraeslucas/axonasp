<%@ Language="VBScript" %>
<%
' Test parser with uploader.asp code in chunks
Response.ContentType = "text/html"
%>
<!DOCTYPE html>
<html>
<head><title>Parser Chunks Test</title></head>
<body>
<h1>Testing Parser with Code Chunks</h1>
<%

' Define a simple class first to verify it works
Response.Write "<h2>Step 1: Define minimal class...</h2>" & vbCrLf
Response.Flush

Dim code1
code1 = "Class TestClass1" & vbCrLf & _
        "  Private x" & vbCrLf & _
        "  Public Sub SetX(v)" & vbCrLf & _
        "    x = v" & vbCrLf & _
        "  End Sub" & vbCrLf & _
        "  Public Function GetX()" & vbCrLf & _
        "    GetX = x" & vbCrLf & _
        "  End Function" & vbCrLf & _
        "End Class"

On Error Resume Next
ExecuteGlobal code1
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ERROR: " & Err.Description & "</p>" & vbCrLf
    Err.Clear
Else
    Response.Write "<p style='color:green'>SUCCESS - TestClass1 defined</p>" & vbCrLf
End If
On Error Goto 0
Response.Flush

' Test nested class with Private Sub Class_Initialize
Response.Write "<h2>Step 2: Class with Class_Initialize...</h2>" & vbCrLf
Response.Flush

Dim code2
code2 = "Class TestClass2" & vbCrLf & _
        "  Private p_val" & vbCrLf & _
        "  Private Sub Class_Initialize()" & vbCrLf & _
        "    p_val = 42" & vbCrLf & _
        "  End Sub" & vbCrLf & _
        "  Private Sub Class_Terminate()" & vbCrLf & _
        "    p_val = 0" & vbCrLf & _
        "  End Sub" & vbCrLf & _
        "  Public Property Get Val()" & vbCrLf & _
        "    Val = p_val" & vbCrLf & _
        "  End Property" & vbCrLf & _
        "End Class"

On Error Resume Next
ExecuteGlobal code2
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ERROR: " & Err.Description & "</p>" & vbCrLf
    Err.Clear
Else
    Response.Write "<p style='color:green'>SUCCESS - TestClass2 defined</p>" & vbCrLf
    Dim t2 : Set t2 = New TestClass2
    Response.Write "<p>TestClass2.Val = " & t2.Val & "</p>" & vbCrLf
    Set t2 = Nothing
End If
On Error Goto 0
Response.Flush

' Test class with more complex features like Property Let
Response.Write "<h2>Step 3: Class with Property Get/Let...</h2>" & vbCrLf
Response.Flush

Dim code3
code3 = "Class TestClass3" & vbCrLf & _
        "  Private p_name" & vbCrLf & _
        "  Public Property Get Name()" & vbCrLf & _
        "    Name = p_name" & vbCrLf & _
        "  End Property" & vbCrLf & _
        "  Public Property Let Name(v)" & vbCrLf & _
        "    p_name = v" & vbCrLf & _
        "  End Property" & vbCrLf & _
        "End Class"

On Error Resume Next
ExecuteGlobal code3
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ERROR: " & Err.Description & "</p>" & vbCrLf
    Err.Clear
Else
    Response.Write "<p style='color:green'>SUCCESS - TestClass3 defined</p>" & vbCrLf
End If
On Error Goto 0
Response.Flush

' Test class with nested Do Loop
Response.Write "<h2>Step 4: Class with Do Loop...</h2>" & vbCrLf
Response.Flush

Dim code4
code4 = "Class TestClass4" & vbCrLf & _
        "  Public Function LoopTest()" & vbCrLf & _
        "    Dim i, result" & vbCrLf & _
        "    i = 0" & vbCrLf & _
        "    result = 0" & vbCrLf & _
        "    Do Until i >= 5" & vbCrLf & _
        "      result = result + i" & vbCrLf & _
        "      i = i + 1" & vbCrLf & _
        "    Loop" & vbCrLf & _
        "    LoopTest = result" & vbCrLf & _
        "  End Function" & vbCrLf & _
        "End Class"

On Error Resume Next
ExecuteGlobal code4
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ERROR: " & Err.Description & "</p>" & vbCrLf
    Err.Clear
Else
    Response.Write "<p style='color:green'>SUCCESS - TestClass4 defined</p>" & vbCrLf
    Dim t4 : Set t4 = New TestClass4
    Response.Write "<p>TestClass4.LoopTest() = " & t4.LoopTest() & "</p>" & vbCrLf
    Set t4 = Nothing
End If
On Error Goto 0
Response.Flush

' Test class with For Each
Response.Write "<h2>Step 5: Class with For Each...</h2>" & vbCrLf
Response.Flush

Dim code5
code5 = "Class TestClass5" & vbCrLf & _
        "  Public Function IterateDict(d)" & vbCrLf & _
        "    Dim k, result" & vbCrLf & _
        "    result = 0" & vbCrLf & _
        "    For Each k In d.Keys" & vbCrLf & _
        "      result = result + 1" & vbCrLf & _
        "    Next" & vbCrLf & _
        "    IterateDict = result" & vbCrLf & _
        "  End Function" & vbCrLf & _
        "End Class"

On Error Resume Next
ExecuteGlobal code5
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ERROR: " & Err.Description & "</p>" & vbCrLf
    Err.Clear
Else
    Response.Write "<p style='color:green'>SUCCESS - TestClass5 defined</p>" & vbCrLf
End If
On Error Goto 0
Response.Flush

' Test multiple classes at once
Response.Write "<h2>Step 6: Two classes at once...</h2>" & vbCrLf
Response.Flush

Dim code6
code6 = "Class ClassA" & vbCrLf & _
        "  Public Name" & vbCrLf & _
        "End Class" & vbCrLf & _
        vbCrLf & _
        "Class ClassB" & vbCrLf & _
        "  Public Value" & vbCrLf & _
        "End Class"

On Error Resume Next
ExecuteGlobal code6
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ERROR: " & Err.Description & "</p>" & vbCrLf
    Err.Clear
Else
    Response.Write "<p style='color:green'>SUCCESS - ClassA and ClassB defined</p>" & vbCrLf
End If
On Error Goto 0
Response.Flush

' Test complex class similar to uploader
Response.Write "<h2>Step 7: Complex class (partial uploader pattern)...</h2>" & vbCrLf
Response.Flush

Dim code7
code7 = "Class TestUploader" & vbCrLf & _
        "  Public UploadedFiles, FormElements, errorMessage" & vbCrLf & _
        "  Private StreamRequest, uploadedYet" & vbCrLf & _
        vbCrLf & _
        "  Private Sub Class_Initialize()" & vbCrLf & _
        "    Set UploadedFiles = Server.CreateObject(""Scripting.Dictionary"")" & vbCrLf & _
        "    Set FormElements = Server.CreateObject(""Scripting.Dictionary"")" & vbCrLf & _
        "    Set StreamRequest = Server.CreateObject(""ADODB.Stream"")" & vbCrLf & _
        "    StreamRequest.Type = 2" & vbCrLf & _
        "    StreamRequest.Open" & vbCrLf & _
        "    uploadedYet = false" & vbCrLf & _
        "  End Sub" & vbCrLf & _
        vbCrLf & _
        "  Private Sub Class_Terminate()" & vbCrLf & _
        "    If IsObject(UploadedFiles) Then" & vbCrLf & _
        "      UploadedFiles.RemoveAll()" & vbCrLf & _
        "      Set UploadedFiles = Nothing" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    If IsObject(FormElements) Then" & vbCrLf & _
        "      FormElements.RemoveAll()" & vbCrLf & _
        "      Set FormElements = Nothing" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    StreamRequest.Close" & vbCrLf & _
        "    Set StreamRequest = Nothing" & vbCrLf & _
        "  End Sub" & vbCrLf & _
        vbCrLf & _
        "  Public Property Get Form(sIndex)" & vbCrLf & _
        "    Form = """"" & vbCrLf & _
        "    If FormElements.Exists(LCase(sIndex)) Then Form = FormElements.Item(LCase(sIndex))" & vbCrLf & _
        "  End Property" & vbCrLf & _
        vbCrLf & _
        "End Class"

On Error Resume Next
ExecuteGlobal code7
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>ERROR: " & Err.Description & "</p>" & vbCrLf
    Err.Clear
Else
    Response.Write "<p style='color:green'>SUCCESS - TestUploader defined</p>" & vbCrLf
    ' Try to instantiate
    Dim tu 
    Set tu = New TestUploader
    Response.Write "<p>TestUploader created successfully!</p>" & vbCrLf
    Set tu = Nothing
End If
On Error Goto 0
Response.Flush

Response.Write "<h2>All tests completed!</h2>" & vbCrLf

%>
</body>
</html>
