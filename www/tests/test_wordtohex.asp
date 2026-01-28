<%
Option Explicit
Response.ContentType = "text/plain"

Class TestWordToHex
    Private m_lOnBits(10)
    
    Public Sub Initialize()
        m_lOnBits(0) = CLng(1)
        m_lOnBits(1) = CLng(3)
        m_lOnBits(2) = CLng(7)
        m_lOnBits(3) = CLng(15)
        m_lOnBits(4) = CLng(31)
        m_lOnBits(5) = CLng(63)
        m_lOnBits(6) = CLng(127)
        m_lOnBits(7) = CLng(255)
        m_lOnBits(8) = CLng(511)
        m_lOnBits(9) = CLng(1023)
        m_lOnBits(10) = CLng(2047)
    End Sub
    
    Private Function RShift(lValue, iShiftBits)
        If iShiftBits = 0 Then
            RShift = lValue
            Exit Function
        End If
        RShift = (lValue And &H7FFFFFFE) \ (2 ^ iShiftBits)
        If (lValue And &H80000000) Then
            RShift = (RShift Or (&H40000000 \ (2 ^ (iShiftBits - 1))))
        End If
    End Function
    
    Public Function WordToHex(lValue)
        Dim lByte, lCount, result
        result = ""
        For lCount = 0 To 3
            lByte = RShift(lValue, lCount * 8) And m_lOnBits(7)
            result = result & Right("0" & Hex(lByte), 2)
        Next
        WordToHex = result
    End Function
    
    Public Function TestIt()
        Initialize()
        Dim testVal
        testVal = 1732584193  ' MD5 initial value for a
        TestIt = WordToHex(testVal)
    End Function
End Class

Dim twth
Set twth = New TestWordToHex
Dim result
result = twth.TestIt()

Response.Write "WordToHex Result: '" & result & "'" & vbCrLf
Response.Write "Length: " & Len(result) & vbCrLf
Response.Write "Expected: '01234567' or similar 8-char hex" & vbCrLf
%>
