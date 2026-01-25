<%
' Simple RegExp test
On Error Resume Next

Dim regex, result
Set regex = Server.CreateObject("G3REGEXP")

If Err.Number <> 0 Then
    Response.Write "ERROR: Cannot create G3REGEXP object: " & Err.Description
Else
    Response.Write "RegExp object created successfully<br><br>"
    
    ' Test 1: Simple pattern matching
    Response.Write "<h3>Test 1: Test Method (Simple Pattern)</h3>"
    result = regex.Test("hello123")
    Response.Write "Test('hello123') with default pattern: " & result & "<br>"
    
    ' Test 2: Set pattern and test
    Response.Write "<h3>Test 2: Test Method with Pattern</h3>"
    result = regex.Test("12345")
    Response.Write "Test('12345') result: " & result & "<br>"
    
    ' Test 3: Replace
    Response.Write "<h3>Test 3: Replace Method</h3>"
    result = regex.Replace("abc123def456", "X")
    Response.Write "Replace result: " & result & "<br>"
    
    ' Test 4: Execute
    Response.Write "<h3>Test 4: Execute Method</h3>"
    Dim matches
    Set matches = regex.Execute("I have 10 apples and 25 oranges")
    If Err.Number = 0 Then
        Response.Write "Execute completed<br>"
        If Not matches Is Nothing Then
            Response.Write "Matches object created<br>"
        End If
    Else
        Response.Write "ERROR in Execute: " & Err.Description & "<br>"
    End If
End If
%>
