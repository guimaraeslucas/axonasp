<% @ Language = "VBScript" %>
<%
Type Point
    X As Integer
    Y As Integer
End Type

Class PointFactory
    Public Function Create(x As Integer, y As Integer) As Point
        Create.X = x
        Create.Y = y
    End Function

    Public Function MovePoint(pt As Point, dx As Integer, dy As Integer) As Point
        MovePoint.X = pt.X + dx
        MovePoint.Y = pt.Y + dy
    End Function
End Class

Dim factory, p1, p2
Set factory = New PointFactory
p1 = factory.Create(10, 20)
p2 = factory.MovePoint(p1, 5, -5)

If p1.X <> 10 Or p1.Y <> 20 Or p2.X <> 15 Or p2.Y <> 15 Then
    Response.Write "FAIL: Class UDT return/params failed (p1=" & p1.X & "," & p1.Y & "|p2=" & p2.X & "," & p2.Y & ")"
    Response.End
End If

Response.Write "PASS: Class UDT return/params are correct"
%>
