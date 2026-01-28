<%
Option Explicit
Response.ContentType = "text/plain"

Class MD5Mini
    Private m_arr(5)
    
    Private Sub ModifyVars(a, b)
        a = a + 1
        b = b + 2
    End Sub
    
    Public Function TestMD5()
        Dim x, y
        x = 10
        y = 20
        
        ModifyVars x, y
        
        TestMD5 = "x=" & x & " y=" & y
    End Function
    
    Public Function TestLoop()
        Dim i, total
        total = 0
        For i = 0 To 5
            m_arr(i) = i * 2
            total = total + m_arr(i)
        Next
        TestLoop = total
    End Function
End Class

Dim mm
Set mm = New MD5Mini

Response.Write "TestMD5: " & mm.TestMD5() & " (expected: x=11 y=22)" & vbCrLf
Response.Write "TestLoop: " & mm.TestLoop() & " (expected: 30)" & vbCrLf
%>
