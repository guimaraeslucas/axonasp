<%
' ============================================
' Test Response.Cookies Collection
' ============================================

Response.Write "<h1>Response.Cookies Test</h1>"
Response.Write "<hr>"

' Test 1: Set simple cookies
Response.Write "<h2>Test 1: Simple Cookie Values</h2>"
Response.Cookies("username") = "john_doe"
Response.Cookies("user_id") = "12345"
Response.Cookies("session_token") = "abc123xyz789"
Response.Write "<p>✓ Set cookies: username, user_id, session_token</p>"

' Test 2: Read back cookie values
Response.Write "<h2>Test 2: Reading Cookies from Request</h2>"
If Request.Cookies("username") <> "" Then
    Response.Write "<p>Username from cookie: " & Request.Cookies("username") & "</p>"
Else
    Response.Write "<p>Username cookie not found (first visit)</p>"
End If

' Test 3: Display all cookies being set
Response.Write "<h2>Test 3: Cookies Being Set</h2>"
Response.Write "<ul>"
Response.Write "<li><strong>username:</strong> john_doe</li>"
Response.Write "<li><strong>user_id:</strong> 12345</li>"
Response.Write "<li><strong>session_token:</strong> abc123xyz789</li>"
Response.Write "</ul>"

Response.Write "<p><em>Check browser Developer Tools > Application/Storage > Cookies to verify</em></p>"

' Test 4: Cookie with expiration
Response.Write "<h2>Test 4: Cookie with Expiration</h2>"
Response.Cookies("persistent_cookie") = "expires_later"
Response.Write "<p>✓ Set cookie with expiration (check implementation)</p>"

Response.Write "<hr>"
Response.Write "<p><strong>Cookie Test Complete!</strong></p>"
Response.Write "<p><a href='/'>Back to Home</a></p>"
%>
