<%
Option Explicit
Response.Write("Testing DEFAULT Property Implementation" & vbCrLf & vbCrLf)

' Simple Class with default property (no arguments)
Class SimpleCounter
    Private Count

    Public Default Property Get Value()
        Value = Count
    End Property

    Public Default Property Let Value(newVal)
        Count = newVal
    End Property

    Public Sub Increment()
        Count = Count + 1
    End Sub
End Class

' Class with indexed default property (using internal array)
Class StringStore
    Private Items()
    Private itemCount

    Public Sub Initialize()
        ReDim Items(9)
        itemCount = 0
    End Sub

    Public Default Property Get Item(index)
        If index >= 0 And index < itemCount Then
            Item = Items(index)
        Else
            Item = Empty
        End If
    End Property

    Public Default Property Let Item(index, val)
        If index >= 0 Then
            If index >= itemCount Then
                itemCount = index + 1
            End If
            Items(index) = val
        End If
    End Property

    Public Function Count()
        Count = itemCount
    End Function
End Class

Response.Write("=" & String(50, "=") & vbCrLf)
Response.Write("1. SIMPLE DEFAULT PROPERTY (No Arguments)      " & vbCrLf)
Response.Write("=" & String(50, "=") & vbCrLf & vbCrLf)

Dim counter
Set counter = New SimpleCounter

' Read via default property; write via explicit property accessor so the
' object variable remains bound to the class instance.
Response.Write("Initial value: " & counter & vbCrLf)
counter.Value = 42
Response.Write("After assignment (counter.Value = 42): " & counter & vbCrLf)

counter.Increment()
Response.Write("After Increment: " & counter & vbCrLf)

counter.Value = counter + 8
Response.Write("After Add (counter.Value = counter + 8): " & counter & vbCrLf & vbCrLf)

Response.Write("=" & String(50, "=") & vbCrLf)
Response.Write("2. INDEXED DEFAULT PROPERTY (With Arguments)    " & vbCrLf)
Response.Write("=" & String(50, "=") & vbCrLf & vbCrLf)

Dim store
Set store = New StringStore
store.Initialize()

' Set via indexed default property
store(0) = "First"
Response.Write("After store(0) = ""First"": store(0) = " & store(0) & vbCrLf)

store(1) = "Second"
Response.Write("After store(1) = ""Second"": store(1) = " & store(1) & vbCrLf)

store(2) = "Third"
Response.Write("After store(2) = ""Third"": store(2) = " & store(2) & vbCrLf)

Response.Write("Total items in store: " & store.Count() & vbCrLf & vbCrLf)

Response.Write("Retrieved values:" & vbCrLf)
Response.Write("  store(0) = " & store(0) & vbCrLf)
Response.Write("  store(1) = " & store(1) & vbCrLf)
Response.Write("  store(2) = " & store(2) & vbCrLf & vbCrLf)

Response.Write("=" & String(50, "=") & vbCrLf)
Response.Write("DEFAULT PROPERTY TESTS COMPLETED SUCCESSFULLY!" & vbCrLf)
Response.Write("=" & String(50, "=") & vbCrLf)
%>
