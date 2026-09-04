# ServerXMLHTTP.Status Property

Returns the HTTP status code from the response.

## Access

Read-only.

## Type

Integer.

- Available after `Send` completes successfully (`ReadyState = 4`).
- Common values: 200 (OK), 404 (Not Found), 500 (Internal Server Error).
- Raises a runtime exception (`0x8000000A`) if read before `Send` has succeeded or after a failed connection attempt.

## Code Example

```asp
<%
Dim oHTTP
Set oHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
oHTTP.Open "GET", "https://example.com/page", False
oHTTP.Send
Select Case oHTTP.Status
    Case 200
        Response.Write "OK: " & oHTTP.ResponseText
    Case 404
        Response.Write "Not found."
    Case Else
        Response.Write "HTTP " & oHTTP.Status & ": " & oHTTP.StatusText
End Select
Set oHTTP = Nothing
%>
```