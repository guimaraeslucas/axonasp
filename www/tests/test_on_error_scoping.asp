<%
' Test: On Error Resume Next per-procedure scoping
' In classic VBScript, On Error Resume Next is scoped per-procedure.
' When a called function has On Error Goto 0 and returns, the
' caller's On Error Resume Next should still be active.

Sub innerSub()
    On Error Resume Next
    Dim x : x = 1 / 0 ' Error, caught
    On Error Goto 0
End Sub

Sub outerSub()
    On Error Resume Next
    innerSub() ' innerSub resets Error, but caller's handler should persist
    Dim y : y = 1 / 0 ' this Error should be caught by outerSub's On Error Resume Next
    Response.Write "outerSub continues after error, err.number=" & Err.number & vbCrLf
    On Error Goto 0
End Sub

outerSub()
Response.Write "TEST OK" & vbCrLf
%>
