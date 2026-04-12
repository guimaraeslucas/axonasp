# ASP Quick Reference

## Intrinsic Objects

- Request: read client input data (query, form, cookies, server variables)
- Response: write output, headers, cookies, redirects
- Server: utility methods and object creation
- Session: per-user state across requests
- Application: global state shared by all requests
- ASPError: structured error information

## Runtime Pattern

1. Read inputs from Request
2. Apply validation and business logic
3. Use libraries through Server.CreateObject
4. Write output through Response
5. Persist state only when needed (Session/Application)

## Minimal ASP Page

```asp
<%
Option Explicit
Dim name
name = Request.QueryString("name")
If Len(Trim(name)) = 0 Then
    name = "Guest"
End If
Response.ContentType = "text/html"
Response.Write "<h1>Hello, " & Server.HTMLEncode(name) & "</h1>"
%>
```
