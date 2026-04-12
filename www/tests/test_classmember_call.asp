<%
' Test 1: Scripting.Dictionary in class method
Class TestClass
    Private m_dict

    Private Sub Class_Initialize()
        Set m_dict = Nothing
    End Sub

    Public Function getDict()
        If m_dict Is Nothing Then
            Set m_dict = Server.CreateObject("scripting.dictionary")
        End If
        Set getDict = m_dict
    End Function

    Public Function testMethod()
        Dim d
        Set d = getDict()
        d.Add "key1", "value1"
        testMethod = d.Count
    End Function
End Class

Dim obj : Set obj = New TestClass
Response.Write "Dict test: " & obj.testMethod() & vbCrLf

' Test 2: Calling zero-arg function via member access in expression
Class TestClass2
    Private p_helper

    Private Sub Class_Initialize()
        Set p_helper = Nothing
    End Sub

    Public Function helper()
        If p_helper Is Nothing Then
            Set p_helper = New HelperClass
        End If
        Set helper = p_helper
    End Function

    Public Function doTest()
        helper.greet("World")
    End Function
End Class

Class HelperClass
    Public Sub greet(name)
        Response.Write "Hello " & name & vbCrLf
    End Sub
End Class

Dim obj2 : Set obj2 = New TestClass2
obj2.doTest()

Response.Write "All tests passed" & vbCrLf
%>
