<% @ Language = "VBScript" %>
<%
Type Point
    X As Integer
    Y As Integer
End Type

Class PointHolder
    Private m_Pt As Point

    Public Sub SetPoint(x As Integer, y As Integer)
        m_Pt.X = x
        m_Pt.Y = y
    End Sub

    Public Function GetX() As Integer
        GetX = m_Pt.X
    End Function

    Public Function GetY() As Integer
        GetY = m_Pt.Y
    End Function
End Class

Dim holder
Set holder = New PointHolder
holder.SetPoint 42, 99

If holder.GetX() <> 42 Or holder.GetY() <> 99 Then
    Response.Write "FAIL: Class UDT field assignment failed (X=" & holder.GetX() & ", Y=" & holder.GetY() & ")"
    Response.End
End If

Response.Write "PASS: Class UDT field assignment is correct"
%>
