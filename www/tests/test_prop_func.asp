<%@ Language="VBScript" %>
<%
Response.ContentType = "text/plain"
Response.Buffer = False
Response.Write "=== TEST SIMPLE CLASS PROPERTY ===" & vbCrLf
Response.Flush

Function helper(s)
    Response.Write "Inside helper(" & s & ")" & vbCrLf
    Response.Flush
    helper = s & "_modified"
End Function

Class SimpleClass
    Private data
    
    Public Property Let Value(v)
        data = v
    End Property
    
    Public Property Get Modified()
        Response.Write "Property Get Modified calling helper..." & vbCrLf
        Response.Flush
        Modified = helper(data)
    End Property
End Class

Dim obj
Set obj = New SimpleClass
obj.Value = "test"
Response.Write "Getting Modified..." & vbCrLf
Response.Flush
Response.Write "Result: " & obj.Modified & vbCrLf
Response.Write "=== DONE ===" & vbCrLf
%>
