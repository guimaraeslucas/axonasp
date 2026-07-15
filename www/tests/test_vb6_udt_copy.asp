<% @ Language = "VBScript" %>
<%
Type Address
    City As String
    Zip As Integer
End Type

Type Person
    Name As String
    Age As Integer
    Home As Address
End Type

' Test 1: Simple UDT assignment
Dim p1 As Person
Dim p2 As Person
p1.Name = "Lucas"
p1.Age = 30
p2 = p1
p2.Name = "G3pix"
p2.Age = 99

If p1.Name <> "Lucas" Or p1.Age <> 30 Or p2.Name <> "G3pix" Or p2.Age <> 99 Then
    Response.Write "FAIL: Simple assignment mutation (p1.Name=" & p1.Name & ", p2.Name=" & p2.Name & ")"
    Response.End
End If

' Test 2: Nested UDT assignment
Dim a As Address
Dim p3 As Person
a.City = "Floripa"
a.Zip = 88000
p3.Home = a
a.City = "Porto"

If p3.Home.City <> "Floripa" Or a.City <> "Porto" Then
    Response.Write "FAIL: Nested assignment mutation (p3.Home.City=" & p3.Home.City & ", a.City=" & a.City & ")"
    Response.End
End If

' Test 3: Array assignment
Dim pts(1) As Address
Dim a2 As Address
a2.City = "Sao Paulo"
pts(0) = a2
a2.City = "Rio"

If pts(0).City <> "Sao Paulo" Or a2.City <> "Rio" Then
    Response.Write "FAIL: Array assignment mutation (pts(0).City=" & pts(0).City & ", a2.City=" & a2.City & ")"
    Response.End
End If

Response.Write "PASS: UDT copy semantics are correct"
%>
