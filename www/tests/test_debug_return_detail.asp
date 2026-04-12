<%
' Debug: Print what is in each register during execution
Class Inner
    Public Value : Sub Class_Initialize() : Value = "init" : End Sub
    End Class

    Class Test
        Function GetInner(cmd)
            Dim obj
            Set obj = New Inner
            obj.Value = cmd
            Response.Write "Before Set: obj=" & obj.Value & vbCrLf
            Set GetInner = obj
            Response.Write "After Set: GetInner=" & GetInner.Value & vbCrLf
        End Function
    End Class

    Dim t
    Set t = New Test

    Dim result
    Result = t.GetInner("test")
%>
