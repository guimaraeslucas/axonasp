<%
' ============================================
' Test Response.End and Response.Redirect
' ============================================

Response.Write "<h1>Testing Response.End</h1>"
Response.Write "<p>This text should appear.</p>"

' Uncomment to test Response.End
' Response.End

Response.Write "<p>This text appears only if Response.End is NOT called.</p>"

' Test Response.Redirect (commented out to allow testing End first)
' Response.Redirect "/test_basics.asp"

Response.Write "<p>End of test.</p>"
%>
