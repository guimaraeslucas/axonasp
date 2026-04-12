<%
' This is a comment with a %> end tag inside it.
' In classic ASP, the block should end above.
' If it works, the text below will be treated as HTML.
%>
<p>This should be literal HTML after the first ASP block.</p>
<%
Response.Write "Second ASP block works if this appears."
%>
<hr>
<%
' Another test with string (escaped to not terminate ASP block):
s = "Some text with %" & "> inside string"
' Or using MS escape if we support it (we don't yet, so use concatenation)
%>
<p>After string with terminator.</p>
<%
Response.Write "Third ASP block works."
%>
