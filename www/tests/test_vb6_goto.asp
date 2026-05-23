<%
' AxonASP Modernization: VB6 GoTo and Labels Test

Sub Assert(condition, message)
    If Not condition Then
        Response.Write "FAIL: " & message & "<br>"
        Response.End
    End If
End Sub

Response.Write "Testing VB6 GoTo and Labels...<br>"

' 1. Basic backward jump (loop)
Dim x
x = 0
LabelLoop:
x = x + 1
If x < 5 Then GoTo LabelLoop
Assert x = 5, "Backward jump failed, expected x=5, got " & x

' 2. Basic forward jump
Dim flag
flag = False
GoTo SkipSetting
flag = True
SkipSetting:
Assert flag = False, "Forward jump failed, flag should be False"

' 3. Numeric labels
Dim y
y = 1
GoTo 100
y = 2
100:
Assert y = 1, "Numeric label jump failed"

' 4. GoTo inside a procedure
Sub TestProc()
    Dim i
    i = 0
    Start:
    i = i + 1
    If i < 3 Then GoTo Start
    Assert i = 3, "GoTo inside Sub failed"
End Sub
TestProc

Response.Write "<b>All GoTo tests passed!</b>"
%>