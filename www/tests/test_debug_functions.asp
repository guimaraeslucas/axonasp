<%
' Test problematic functions

Response.Write "Testing problematic functions:<br><br>"

' Test 1: Document.Write
Response.Write "1. Document.Write test:<br>"
Response.Write "<pre>"
Dim htmlContent
htmlContent = "<script>alert('XSS')</script>ddd"
Response.Write "Before: " & htmlContent & vbCrLf
Response.Write "After Document.Write: "
Document.Write htmlContent
Response.Write "</pre><br>"

' Test 2: AxTime
Response.Write "2. AxTime() test:<br>"
Response.Write "<pre>"
On Error Resume Next
Dim timestamp
timestamp = AxTime()
Response.Write "AxTime() = " & timestamp
If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Description
End If
On Error Goto 0
Response.Write "</pre><br>"

' Test 3: AxGenerateGuid
Response.Write "3. AxGenerateGuid() test:<br>"
Response.Write "<pre>"
On Error Resume Next
Dim guid
guid = AxGenerateGuid()
Response.Write "AxGenerateGuid() = " & guid
If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Description
End If
On Error Goto 0
Response.Write "</pre><br>"

' Test 4: AxBuildQueryString
Response.Write "4. AxBuildQueryString() test:<br>"
Response.Write "<pre>"
On Error Resume Next
Dim params, queryString
Set params = CreateObject("Scripting.Dictionary")
params("name") = "John"
params("age") = 25
params("city") = "New York"
queryString = AxBuildQueryString(params)
Response.Write "AxBuildQueryString result: " & queryString
If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Description
End If
On Error Goto 0
Response.Write "</pre><br>"

' Test 5: AxGetRequest
Response.Write "5. AxGetRequest() test:<br>"
Response.Write "<pre>"
On Error Resume Next
Dim reqParams
reqParams = AxGetRequest()
If IsObject(reqParams) Then
    Response.Write "AxGetRequest returned: " & TypeName(reqParams)
    If reqParams.Count > 0 Then
        Dim key
        For Each key In reqParams.Keys
            Response.Write vbCrLf & "  " & key & " = " & reqParams(key)
        Next
    Else
        Response.Write "(empty)"
    End If
Else
    Response.Write "AxGetRequest returned: " & TypeName(reqParams) & " = " & reqParams
End If
If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Description
End If
On Error Goto 0
Response.Write "</pre><br>"
%>
