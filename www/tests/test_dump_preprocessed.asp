<%
' Test the preprocessed source dump feature
Response.Write "Before include" & vbCrLf
%>
<!-- #include file="include_test.inc" -->
<%
Response.Write "After include" & vbCrLf
%>
