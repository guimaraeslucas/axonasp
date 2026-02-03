<%@ Language="VBScript" %>
<%
Response.ContentType = "text/plain"
Response.Buffer = False
Response.Write "=== TEST PROPERTY GET IN PROPERTY GET ===" & vbCrLf
Response.Flush

Function helper(s)
    Response.Write "Inside helper(" & s & ")" & vbCrLf
    Response.Flush
    helper = s & "_modified"
End Function

Class TestClass
    Private data
    
    Public Property Let Value(v)
        data = v
    End Property
    
    Public Property Get Value()
        Response.Write "Property Get Value returning " & data & vbCrLf
        Response.Flush
        Value = data
    End Property
    
    Public Property Get Modified()
        Response.Write "Property Get Modified calling helper(Value)..." & vbCrLf
        Response.Flush
        Modified = helper(Value)  ' Note: Value is another Property Get!
        Response.Write "Modified result..." & vbCrLf
        Response.Flush
    End Property
End Class

Response.Write "Creating object..." & vbCrLf
Response.Flush
Dim obj
Set obj = New TestClass
obj.Value = "test"
Response.Write "Getting Modified..." & vbCrLf
Response.Flush
Response.Write "Result: " & obj.Modified & vbCrLf
Response.Write "=== DONE ===" & vbCrLf
%>
