<%
Option Explicit
Response.Write "Testing Array Returns from Class Properties" & vbCrLf & vbCrLf

' Test 1: Simple REDIM with property
Class SimpleArray
    Private m_items

    Private Sub Class_Initialize()
        ReDim m_items( - 1)
    End Sub

    Public Property Get Items()
        Items = m_items
    End Property

    Public Sub AddItem(val)
        ReDim Preserve m_items(UBound(m_items) + 1)
        m_items(UBound(m_items)) = val
    End Sub
End Class

Response.Write "Test 1: Simple Array Property" & vbCrLf
Dim obj1
Set obj1 = New SimpleArray
obj1.AddItem "first"
obj1.AddItem "second"

Dim result1
result1 = obj1.Items
Response.Write "TypeName(result1): " & TypeName(result1) & vbCrLf
Response.Write "UBound(result1): " & UBound(result1) & vbCrLf
Response.Write "result1(0): " & result1(0) & vbCrLf
Response.Write "result1(1): " & result1(1) & vbCrLf

' Test 2: Abstract array return (REDIM to -1)
Class EmptyArrayTest
    Private m_items

    Private Sub Class_Initialize()
        ReDim m_items( - 1)
    End Sub

    Public Property Get Items()
        Items = m_items
    End Property
End Class

Response.Write vbCrLf & "Test 2: Empty Array Property (ReDim m_items(-1))" & vbCrLf
Dim obj2
Set obj2 = New EmptyArrayTest

Dim result2
result2 = obj2.Items
Response.Write "TypeName(result2): " & TypeName(result2) & vbCrLf

' Test 3: ReDim Preserve in property getter
Class ReDimPreserveTest
    Private m_items
    Private m_count

    Private Sub Class_Initialize()
        ReDim m_items( - 1)
        m_count = 0
    End Sub

    Public Sub AddItem(val)
        ReDim Preserve m_items(UBound(m_items) + 1)
        m_items(UBound(m_items)) = val
        m_count = m_count + 1
    End Sub

    Public Property Get Items()
        Dim tmp
        tmp = m_items
        If m_count > 0 Then
            ReDim Preserve tmp(m_count - 1)
        End If
        Items = tmp
    End Property
End Class

Response.Write vbCrLf & "Test 3: ReDim Preserve in Property Getter" & vbCrLf
Dim obj3
Set obj3 = New ReDimPreserveTest
obj3.AddItem "A"
obj3.AddItem "B"
obj3.AddItem "C"

Dim result3
result3 = obj3.Items
Response.Write "TypeName(result3): " & TypeName(result3) & vbCrLf
Response.Write "UBound(result3): " & UBound(result3) & vbCrLf
Response.Write "result3(0): " & result3(0) & vbCrLf
Response.Write "result3(1): " & result3(1) & vbCrLf
Response.Write "result3(2): " & result3(2) & vbCrLf

Response.Write vbCrLf & "ALL ARRAY TESTS COMPLETED" & vbCrLf
%>
