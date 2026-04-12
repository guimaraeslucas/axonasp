<%
Response.Write("Testing Property Initialization" & vbCrLf)

Class SimpleProperty
    Private m_value

    Public Sub Initialize(v)
        Response.Write("Initialize called with: " & v & vbCrLf)
        m_value = v
    End Sub

    Public Property Get Value()
        Response.Write("Property Get Value called" & vbCrLf)
        Value = m_value
    End Property

    Public Property Let Value(v)
        Response.Write("Property Let Value called with: " & v & vbCrLf)
        m_value = v
    End Property
End Class

Response.Write("Creating object..." & vbCrLf)
Dim obj
Set obj = New SimpleProperty

Response.Write("Calling Initialize..." & vbCrLf)
obj.Initialize(42)

Response.Write("Getting value..." & vbCrLf)
Response.Write("Value: " & obj.Value & vbCrLf)

Response.Write("Setting value to 99..." & vbCrLf)
obj.Value = 99

Response.Write("Getting value again..." & vbCrLf)
Response.Write("Value: " & obj.Value & vbCrLf)
%>
