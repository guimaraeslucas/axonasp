<%
' Teste melhorado para class array properties
' Verifica se o array retornado é realmente um array e suas propriedades

Class TestArray
    Dim i_items

    Private Sub Class_Initialize()
        ReDim i_items( - 1)
    End Sub

    Public Property Get Items()
        Dim tmp
        tmp = i_items
        Items = tmp
    End Property

    Public Sub AddItem(val)
        ReDim Preserve i_items(UBound(i_items) + 1)
        i_items(UBound(i_items)) = val
    End Sub
End Class

Response.Write "<h2>Test 1: Empty Array</h2>"
Dim obj1
Set obj1 = New TestArray
Dim result1
result1 = obj1.Items
Response.Write "TypeName(result1)=" & TypeName(result1) & "<br>"
Response.Write "IsArray(result1)=" & IsArray(result1) & "<br>"
Response.Write "UBound(result1)=" & UBound(result1) & "<br>"

Response.Write "<h2>Test 2: Array with Items</h2>"
Dim obj2
Set obj2 = New TestArray
obj2.AddItem "First"
obj2.AddItem "Second"
obj2.AddItem "Third"
Dim result2
result2 = obj2.Items
Response.Write "TypeName(result2)=" & TypeName(result2) & "<br>"
Response.Write "IsArray(result2)=" & IsArray(result2) & "<br>"
Response.Write "UBound(result2)=" & UBound(result2) & "<br>"
Response.Write "LBound(result2)=" & LBound(result2) & "<br>"
Response.Write "Item count: " & (UBound(result2) - LBound(result2) + 1) & "<br>"
Response.Write "First item=" & result2(0) & "<br>"
Response.Write "Last item=" & result2(UBound(result2)) & "<br>"
%>
