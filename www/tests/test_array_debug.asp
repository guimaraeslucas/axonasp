<%
debug_asp_code = "FALSE"

' Simple array test to understand array behavior in VM

' Test 1: Direct array assignment
Response.Write "Test 1: Direct array assignment" & vbCrLf
Dim arr1
ReDim arr1( - 1)
Response.Write "After ReDim(-1): TypeName=" & TypeName(arr1) & ", IsArray=" & IsArray(arr1) & vbCrLf

ReDim Preserve arr1(2)
arr1(0) = "A"
arr1(1) = "B"
arr1(2) = "C"
Response.Write "After ReDim(2) and fill: TypeName=" & TypeName(arr1) & ", IsArray=" & IsArray(arr1) & vbCrLf
Response.Write "  UBound(arr1)=" & UBound(arr1) & vbCrLf

' Test 2: Array copy by assignment
Response.Write vbCrLf & "Test 2: Array copy by assignment" & vbCrLf
Dim arr2Copy
arr2Copy = arr1
Response.Write "After arr2Copy = arr1: TypeName=" & TypeName(arr2Copy) & ", IsArray=" & IsArray(arr2Copy) & vbCrLf
Response.Write "  UBound(arr2Copy)=" & UBound(arr2Copy) & vbCrLf

' Test 3: Array in function
Response.Write vbCrLf & "Test 3: Array in class property" & vbCrLf
Class SimpleArrayClass
    Private m_arr

    Private Sub Class_Initialize()
        ReDim m_arr( - 1)
    End Sub

    Public Property Get MyArray()
        Dim result
        result = m_arr
        MyArray = result
    End Property

    Public Sub AddItem(val)
        ReDim Preserve m_arr(UBound(m_arr) + 1)
        m_arr(UBound(m_arr)) = val
    End Sub
End Class

Dim obj
Set obj = New SimpleArrayClass
obj.AddItem "First"
obj.AddItem "Second"

Dim arrFromProperty
arrFromProperty = obj.MyArray
Response.Write "From property: TypeName=" & TypeName(arrFromProperty) & ", IsArray=" & IsArray(arrFromProperty) & vbCrLf
Response.Write "  UBound=" & UBound(arrFromProperty) & vbCrLf

' Test 4: Check if both reference same object
Response.Write vbCrLf & "Test 4: Array reference semantics" & vbCrLf
Dim arr3
ReDim arr3(1)
arr3(0) = "Original"
arr3(1) = "Val"

Dim arr3Copy
arr3Copy = arr3
arr3Copy(0) = "Modified"
Response.Write "Original arr3(0)=" & arr3(0) & vbCrLf
Response.Write "Should be 'Modified' if pass-by-reference: " & (arr3(0) = "Modified") & vbCrLf
%>
