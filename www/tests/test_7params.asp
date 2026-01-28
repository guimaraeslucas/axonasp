<%
Option Explicit
Response.ContentType = "text/plain"

Class MultiParam
    Private Sub Modify7(a, b, c, d, e, f, g)
        a = a + 1
        b = b + 2
        c = c + 3
        d = d + 4
        e = e + 5
        f = f + 6
        g = g + 7
    End Sub
    
    Public Function Test7Params()
        Dim v1, v2, v3, v4, v5, v6, v7
        v1 = 1
        v2 = 2
        v3 = 3
        v4 = 4
        v5 = 5
        v6 = 6
        v7 = 7
        
        Modify7 v1, v2, v3, v4, v5, v6, v7
        
        Test7Params = v1 & "," & v2 & "," & v3 & "," & v4 & "," & v5 & "," & v6 & "," & v7
    End Function
End Class

Dim mp
Set mp = New MultiParam
Response.Write "Result: " & mp.Test7Params() & vbCrLf
Response.Write "Expected: 2,4,6,8,10,12,14"
%>
