<!-- #include file="asplite/asplite.asp"-->
<%
Response.Write "Testing aspl availability<br><br>"

Response.Write "aspl is defined: " & (not (aspl is nothing)) & "<br>"
Response.Write "aspl typename: " & typename(aspl) & "<br><br>"

' Now test if aspl is available after executeGlobal
dim testCode
testCode = "Response.Write ""Inside executeGlobal: aspl is "" & (not (aspl is nothing)) & ""<br>"""

executeGlobal testCode
%>
