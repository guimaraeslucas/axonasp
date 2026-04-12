<%
Response.Write("Testing Indexed Property (Array Property)" & vbCrLf & vbCrLf)

Class SimpleArray
    Private m_items()
    Private m_count

    Public Sub Initialize()
        ReDim m_items(9)
        m_count = 0
        Response.Write("Array initialized" & vbCrLf)
    End Sub

    Public Property Get Item(index)
        Response.Write("Get Item(" & index & ") called" & vbCrLf)
        If index >  = 0 And index < m_count Then
            Item = m_items(index)
        Else
            Item = Empty
        End If
    End Property

    Public Property Let Item(index, value)
        Response.Write("Set Item(" & index & ") = " & value & vbCrLf)
        If index >  = m_count Then
            m_count = index + 1
            If index > UBound(m_items) Then
                ReDim Preserve m_items(index + 5)
            End If
        End If
        m_items(index) = value
    End Property

    Public Function Count()
        Count = m_count
    End Function
End Class

Dim arr
Set arr = New SimpleArray
arr.Initialize()

Response.Write("Setting items..." & vbCrLf)
arr.Item(0) = "First"
arr.Item(1) = "Second"
arr.Item(2) = "Third"

Response.Write(vbCrLf & "Getting items..." & vbCrLf)
Response.Write("Item(0): " & arr.Item(0) & vbCrLf)
Response.Write("Item(1): " & arr.Item(1) & vbCrLf)
Response.Write("Item(2): " & arr.Item(2) & vbCrLf)
Response.Write("Count: " & arr.Count() & vbCrLf)
%>
