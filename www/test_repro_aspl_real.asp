<!-- #include file="aspLite-master/aspLite/aspLite.asp"-->
<%
Response.Write "<h3>Starting Real aspl Test</h3>"

Response.Write "<h4>1. Testing aspl.isEmpty(""foo"")</h4>"
Dim res
res = aspl.isEmpty("foo")
Response.Write "Result: " & res & " (Expected: False)<br>"
If res <> False Then Response.Write "FAIL: aspl.isEmpty(""foo"") returned True<br>"

Response.Write "<h4>2. Testing aspl.loadText(""test_repro_dummy.inc"")</h4>"
Dim content
content = aspl.loadText("test_repro_dummy.inc")
Response.Write "Content length: " & Len(content) & "<br>"
If Len(content) > 0 Then 
    Response.Write "PASS: loadText worked<br>" 
Else 
    Response.Write "FAIL: loadText returned empty (Stream issue?)<br>"
    Response.Write "Last Error: " & aspl.asperror("loadText check") & "<br>"
End If

Response.Write "<h4>3. Testing aspl(""test_repro_dummy.inc"") (Default Method)</h4>"
On Error Resume Next
aspl("test_repro_dummy.inc")
If Err.Number <> 0 Then 
    Response.Write "FAIL: Error calling aspl(...): " & Err.Description & "<br>"
Else
    Response.Write "Call finished (check for PASS output above/below)<br>"
End If
On Error Goto 0

Response.Write "<h3>End Real aspl Test</h3>"
%>
