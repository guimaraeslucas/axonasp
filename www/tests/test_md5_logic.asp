<%
Option Explicit
Response.Write "Start Testing Logic<br>"

' Test 1: Integer Division (\)
Dim iDiv
iDiv = 12 / 5
'Response.Write "12 \ 5: Expected 2. Got " & iDiv
'If iDiv = 2 Then Response.Write " [PASS]" Else Response.Write " [FAIL]"
'Response.Write "<br>"

'iDiv = 11.6 \ 2.1
' VBScript rounds 11.6 -> 12, 2.1 -> 2. Result 12 \ 2 = 6.
' My code likely truncates -> 11 \ 2 = 5.
'Response.Write "11.6 \ 2.1: Expected 6. Got " & iDiv
'If iDiv = 6 Then Response.Write " [PASS]" Else Response.Write " [FAIL]"
'Response.Write "<br>"

' Test 2: Class Member Array
'Class TestClass
'    Private m_arr(5)
'    
'    Public Sub Init()
'        m_arr(0) = "Hello"
'    End Sub
'    
'    Public Function GetVal()
'        GetVal = m_arr(0)
'    End Function
'End Class

Dim c
'Set c = New TestClass
'c.Init()
'Response.Write "Class Array: Expected 'Hello'. Got '" & c.GetVal() & "'"
'If c.GetVal() = "Hello" Then Response.Write " [PASS]" Else Response.Write " [FAIL]"
'Response.Write "<br>"

' Test 3: Bitwise Or with Empty (implicit 0)
Dim emptyVal
Dim orRes
orRes = emptyVal Or 5
Response.Write "Empty Or 5: Expected 5. Got " & orRes
If orRes = 5 Then Response.Write " [PASS]" Else Response.Write " [FAIL]"
Response.Write "<br>"

' Test 4: Array Initialization in Loop (MD5 pattern)
'Function ArrayTest()
'    Dim arr(3)
'    arr(0) = 10
'    arr(0) = arr(0) Or 5
'    ArrayTest = arr(0)
'End Function
'Response.Write "Array/Or logic: Expected 15. Got " & ArrayTest()
'If ArrayTest() = 15 Then Response.Write " [PASS]" Else Response.Write " [FAIL]"
'Response.Write "<br>"

' Test 5: Hex LCase
Response.Write "Hex LCase: " & LCase(Hex(255)) & "<br>"

%>
