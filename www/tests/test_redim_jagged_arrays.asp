<%
Dim outer()
Dim inner(2)
inner(0) = "a"
inner(1) = "b"
inner(2) = "c"

ReDim outer(0)
outer(0) = inner

' ReDim Preserve must only evaluate the outer array rank.
ReDim Preserve outer(1)
outer(1) = "1"

Response.Write outer(0)(0) & "," & outer(0)(1) & "," & outer(0)(2) & "<br>"
Response.Write outer(1)
%>