<%@ Language=VBScript %>
<%
Class Singleton
    Private m_Name

    Private Sub Class_Initialize
        m_Name = "The Only One"
    End Sub

    Public Property Get Name
        Name = m_Name
    End Property
End Class

Function GetInstance()
    Static instance
    If instance Is Nothing Then
        Set instance = New Singleton
    End If
    Set GetInstance = instance
End Function

Dim s1, s2
Set s1 = GetInstance()
Set s2 = GetInstance()

Response.Write "s1.Name = " & s1.Name & vbCrLf
Response.Write "s2.Name = " & s2.Name & vbCrLf

If s1 Is s2 Then
    Response.Write "Same object — Static singleton works" & vbCrLf
Else
    Response.Write "Different objects — Static singleton failed" & vbCrLf
End If
%>