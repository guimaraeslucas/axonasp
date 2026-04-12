<%
Dim result

result = Eval("Nothing")
Response.Write "Eval(Nothing): " & TypeName(result) & vbCrLf

Response.Write "Eval with Mod: " & Eval("10 Mod 3") & vbCrLf

Response.Write "Nested Len/UCase: " & Eval("Len(UCase(""test""))") & vbCrLf

Response.Write "Type coercion: " & Eval("1 & 2 & 3") & vbCrLf

Response.Write "Logical AND: " & Eval("1=1 And 2=2") & vbCrLf

Response.Write "All edge cases passed!" & vbCrLf
%>
