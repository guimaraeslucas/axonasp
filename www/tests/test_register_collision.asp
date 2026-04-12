<%
' Test for Register Collision Fix
' This tests if recursive method calls properly preserve object references

Class TestObject
    Public Function Execute(sql)
        ' Simulate database execution
        Dim result
        result = "SELECT * FROM test"
        Execute = result
    End Function
End Class

Dim obj
Set obj = New TestObject

' This should work without shadowing obj with a boolean
Dim result
result = obj.Execute("test sql")

Response.Write "Result: " & result & vbCrLf

' Test with nested calls
Function CallMethod(o)
    ' Comparison that returns boolean - could shadow 'o' if register allocation is bad
    If 1 > 0 Then
        CallMethod = o.Execute("nested")
    Else
        CallMethod = "error"
    End If
End Function

Dim result2
result2 = CallMethod(obj)
Response.Write "Result2: " & result2 & vbCrLf

' Test with default properties
Class TestDefault
    Public Default Function Item(Key)
        Item = "value_" & Key
    End Function
End Class

Dim def
Set def = New TestDefault
Response.Write "Default: " & def("test") & vbCrLf
%>
