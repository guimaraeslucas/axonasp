<%
' Test Issues 12, 13, 14 - DETAILED

Response.Write "<h2>Issue 12: Err Object - Complete Test</h2>"
Response.Write "<pre>"

On Error Resume Next

Set Err_Test = New ErrTestClass
Err_Test.RaiseError
If Err.Number <> 0 Then
    Response.Write "Err.Number: " & Err.Number & vbCrLf
    Response.Write "Err.Description: " & Err.Description & vbCrLf
    Response.Write "Err.Source: " & Err.Source & vbCrLf
    Response.Write "Err.ASPCode: " & Err.ASPCode & vbCrLf
    Response.Write "Err.Category: " & Err.Category & vbCrLf
    Response.Write "Err.File: " & Err.File & vbCrLf
    Response.Write "Err.Line: " & Err.Line & vbCrLf
    Response.Write "Err.Column: " & Err.Column & vbCrLf
    Response.Write "Err.ASPDescription: " & Err.ASPDescription & vbCrLf
    Err.Clear
End If

Response.Write "</pre>"

Response.Write "<h2>Issue 13: RegExp - SubMatches Complete Test</h2>"
Response.Write "<pre>"

Set regex = New RegExp
regex.Pattern = "(\d+)-(\w+)-([a-z]+)"
regex.Global = False
regex.IgnoreCase = False

Dim testStr, matches, match, subMatches, i

testStr = "123-ABC-test"
Set matches = regex.Execute(testStr)

Response.Write "Test String: " & testStr & vbCrLf
Response.Write "Pattern: " & regex.Pattern & vbCrLf
Response.Write "Total Matches: " & matches.Count & vbCrLf

If matches.Count > 0 Then
    Set match = matches.Item(0)
    Response.Write vbCrLf & "First Match:" & vbCrLf
    Response.Write "  Value: " & match.Value & vbCrLf
    Response.Write "  Index: " & match.Index & vbCrLf
    Response.Write "  Length: " & match.Length & vbCrLf

    If Not IsEmpty(match.SubMatches) Then
        Set subMatches = match.SubMatches
        Response.Write "  SubMatches Count: " & subMatches.Count & vbCrLf
        If subMatches.Count > 0 Then
            For i = 0 To subMatches.Count - 1
                Response.Write "    SubMatch[" & i & "].Value: " & subMatches.Item(i).Value & vbCrLf
            Next
        End If
    Else
        Response.Write "  SubMatches: EMPTY" & vbCrLf
    End If
End If

Response.Write "</pre>"

Response.Write "<h2>Issue 14: Dictionary - Default Property Test</h2>"
Response.Write "<pre>"

Set dict = Server.CreateObject("Scripting.Dictionary")

' Test 1: Traditional method
dict.Add "test1", "value1"
Response.Write "Dict.Add 'test1', 'value1'" & vbCrLf
Response.Write "Dict('test1'): " & dict("test1") & vbCrLf
Response.Write "Dict.Item('test1'): " & dict.Item("test1") & vbCrLf

' Test 2: Direct indexing assignment
dict("test2") = "value2"
Response.Write vbCrLf & "Dict('test2') = 'value2'" & vbCrLf
Response.Write "Dict('test2'): " & dict("test2") & vbCrLf

' Test 3: Mixed assignment and access
Dim Key, value
For Each Key In dict
    Response.Write "Dict('" & Key & "'): " & dict(Key) & vbCrLf
Next

Response.Write vbCrLf & "Count: " & dict.Count & vbCrLf

Response.Write "</pre>"

Class ErrTestClass
    Public Sub RaiseError()
        Err.Raise 999, "TestClass", "This is a test error"
    End Sub
End Class
%>
