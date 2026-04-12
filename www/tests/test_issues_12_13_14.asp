<%
' Test Issues 12, 13, 14 from AXONASP_AUDIT

Response.Write "<h2>Testing Issues 12, 13, 14</h2>"

' ========================================
' ISSUE 12: ASPError Object Completeness
' ========================================
Response.Write "<h3>Issue 12: Err Object Properties</h3>"
Response.Write "<pre>"

On Error Resume Next

Err.Clear
Response.Write "- Err.Number: " & Err.Number & vbCrLf
Response.Write "- Err.Description: '" & Err.Description & "'" & vbCrLf
Response.Write "- Err.Source: '" & Err.Source & "'" & vbCrLf

' Try to raise an error
Err.Raise 11, "TestSource", "Test Description"
Response.Write "After Raise:" & vbCrLf
Response.Write "- Err.Number: " & Err.Number & vbCrLf
Response.Write "- Err.Description: '" & Err.Description & "'" & vbCrLf
Response.Write "- Err.Source: '" & Err.Source & "'" & vbCrLf

' Check if these properties exist (extended ASPError properties)
On Error Resume Next
Response.Write "- Err.ASPCode exists: " & HasProperty(Err, "ASPCode") & vbCrLf
Response.Write "- Err.Category exists: " & HasProperty(Err, "Category") & vbCrLf
Response.Write "- Err.File exists: " & HasProperty(Err, "File") & vbCrLf
Response.Write "- Err.Line exists: " & HasProperty(Err, "Line") & vbCrLf
Response.Write "- Err.Column exists: " & HasProperty(Err, "Column") & vbCrLf
Response.Write "- Err.ASPDescription exists: " & HasProperty(Err, "ASPDescription") & vbCrLf

Response.Write "</pre>"

' ========================================
' ISSUE 13: RegExp Object Return Values
' ========================================
Response.Write "<h3>Issue 13: RegExp Execute() Return Values</h3>"
Response.Write "<pre>"

Set regex = New RegExp
regex.Pattern = "(\w+)@(\w+)\.(\w+)"
regex.Global = True
regex.IgnoreCase = True

Dim testEmail
testEmail = "test@example.com another@domain.org"

Set matches = regex.Execute(testEmail)

Response.Write "Pattern: " & regex.Pattern & vbCrLf
Response.Write "Global: " & regex.Global & vbCrLf
Response.Write "IgnoreCase: " & regex.IgnoreCase & vbCrLf
Response.Write "Matches Count: " & matches.Count & vbCrLf
Response.Write vbCrLf

If matches.Count > 0 Then
    Response.Write "First Match Properties:" & vbCrLf
    Set match = matches.Item(0)
    Response.Write "- match.Value: '" & match.Value & "'" & vbCrLf
    Response.Write "- match.FirstIndex: " & match.FirstIndex & vbCrLf
    Response.Write "- match.Index: " & match.Index & vbCrLf
    Response.Write "- match.Length: " & match.Length & vbCrLf
    Response.Write "- match.SubMatches exists: " & HasProperty(match, "SubMatches") & vbCrLf

    If HasProperty(match, "SubMatches") Then
        Set subMatches = match.SubMatches
        Response.Write "- SubMatches Count: " & subMatches.Count & vbCrLf
    End If
End If

Response.Write "</pre>"

' ========================================
' ISSUE 14: Dictionary Default Property
' ========================================
Response.Write "<h3>Issue 14: Dictionary Default Property (Item)</h3>"
Response.Write "<pre>"

Set dict = CreateObject("Scripting.Dictionary")

' Test traditional Item method
dict.Item("key1") = "value1"
Response.Write "Using dict.Item('key1') = 'value1'" & vbCrLf
Response.Write "Result: " & dict.Item("key1") & vbCrLf
Response.Write vbCrLf

' Test default property syntax (without Item)
On Error Resume Next
dict("key2") = "value2"
If Err.Number = 0 Then
    Response.Write "Using dict('key2') = 'value2' [SUCCESS]" & vbCrLf
    Response.Write "Result: " & dict("key2") & vbCrLf
Else
    Response.Write "Using dict('key2') = 'value2' [FAILED]" & vbCrLf
    Response.Write "Error: " & Err.Description & vbCrLf
End If
On Error Goto 0

Response.Write "</pre>"

Function HasProperty(obj, propName)
    On Error Resume Next
    Dim dummy
    dummy = obj.GetProperty(propName)
    If Err.Number = 0 Then
        HasProperty = True
    Else
        HasProperty = False
    End If
    Err.Clear
End Function
%>
