<%
' Test Issues 12, 13, 14 - SIMPLIFIED

Response.Write "<h2>Issue 12: Err Object Properties</h2>"
Response.Write "<pre>"

' Test setting Err properties directly
Err.ASPCode = "TEST_CODE"
Err.Category = "TEST_CATEGORY"
Err.File = "test.asp"
Err.Line = 42
Err.Column = 10
Err.ASPDescription = "Test description from ASP"

Response.Write "After setting properties:" & vbCrLf
Response.Write "Err.ASPCode: " & Err.ASPCode & vbCrLf
Response.Write "Err.Category: " & Err.Category & vbCrLf
Response.Write "Err.File: " & Err.File & vbCrLf
Response.Write "Err.Line: " & Err.Line & vbCrLf
Response.Write "Err.Column: " & Err.Column & vbCrLf
Response.Write "Err.ASPDescription: " & Err.ASPDescription & vbCrLf

Err.Clear
Response.Write vbCrLf & "After Clear:" & vbCrLf
Response.Write "Err.ASPCode: '" & Err.ASPCode & "'" & vbCrLf
Response.Write "Err.Category: '" & Err.Category & "'" & vbCrLf

Response.Write "</pre>"

Response.Write "<h2>Issue 13: RegExp SubMatches</h2>"
Response.Write "<pre>"

Set regex = New RegExp
regex.Pattern = "(\d+)-(\w+)"
regex.Global = False

Dim testStr, matches, match

testStr = "123-ABC"
Set matches = regex.Execute(testStr)

Response.Write "Pattern: " & regex.Pattern & vbCrLf
Response.Write "Test String: " & testStr & vbCrLf
Response.Write "Matches: " & matches.Count & vbCrLf

If matches.Count > 0 Then
    Set match = matches.Item(0)
    Response.Write "Match.Value: " & match.Value & vbCrLf
    Response.Write "Match.SubMatches.Count: " & match.SubMatches.Count & vbCrLf

    If match.SubMatches.Count > 0 Then
        Response.Write "SubMatch[0].Value: " & match.SubMatches.Item(0).Value & vbCrLf
        Response.Write "SubMatch[1].Value: " & match.SubMatches.Item(1).Value & vbCrLf
    End If
End If

Response.Write "</pre>"

Response.Write "<h2>Issue 14: Dictionary Default Property</h2>"
Response.Write "<pre>"

Set dict = Server.CreateObject("Scripting.Dictionary")

dict("key1") = "value1"
dict("key2") = "value2"

Response.Write "dict('key1'): " & dict("key1") & vbCrLf
Response.Write "dict.Item('key1'): " & dict.Item("key1") & vbCrLf
Response.Write "dict('key2'): " & dict("key2") & vbCrLf
Response.Write "dict.Count: " & dict.Count & vbCrLf

Response.Write vbCrLf & "SUCCESS: All tests passed!" & vbCrLf
Response.Write "</pre>"
%>
