<%@ Language=VBScript %>
<%
Response.ContentType = "text/plain"

' ─── Test 1: Basic On Error Resume Next ───────────────────────
On Error Resume Next
Dim errVal
errVal = 1 / 0
Dim caught
caught = (Err.Number <> 0)
On Error GoTo 0

If caught Then
    Response.Write("[PASS] On Error Resume Next (caught div by zero)" & vbCrLf)
Else
    Response.Write("[FAIL] On Error Resume Next (did not catch)" & vbCrLf)
End If

Response.Write("RESULT: TEST_DONE" & vbCrLf)
%>