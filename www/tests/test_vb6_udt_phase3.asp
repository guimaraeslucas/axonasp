<% @ Language = "VBScript" %>
<%
Type Address
    City As String
End Type

Type Person
    Name As String
    Age As Integer
    Home As Address
End Type

Dim addr As Address
Dim p As Person

addr.City = "Sao Paulo"
p.Name = "Maya"
p.Age = 29
p.Home = addr

Dim actual
actual = p.Name & "|" & p.Age & "|" & p.Home.City
Response.Write "PASS:VB6_UDT_PHASE3|" & actual
%>
