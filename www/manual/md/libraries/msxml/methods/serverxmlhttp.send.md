# ServerXMLHTTP.Send Method

Executes the HTTP request configured by `Open`. Blocks until the response is received or the timeout expires.

## Syntax

```asp
objHTTP.Send [body]
```

## Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `body` | String or Byte Array | No | The request body to send. For `POST`/`PUT` requests, supply the payload here. String values are encoded as UTF-8. Byte arrays are sent as-is. |

## Return Value

Empty. This method does not return a value. Read `ResponseText`, `ResponseBody`, or `Status` after `Send` completes.

## Remarks

- If no `Content-Type` header is set and a String body is provided, `application/x-www-form-urlencoded` is used automatically.
- If no `Content-Type` header is set and a byte array body is provided, `application/octet-stream` is used automatically.
- A default `User-Agent` header is added automatically if none is set.
- The request timeout defaults to 30 seconds and can be changed via the `Timeout` property before calling `Send`.
- On completion, `ReadyState` is set to 4.
- Method names are case-insensitive.
- Calling `Send` after an invalid or missing `Open` raises error `0x80004005` ("Unspecified error").
- Network transport failures (such as connection refused, DNS resolution failure, or timeout) raise error `0x80072EFD` ("A connection with the server could not be established"). Internal transport error messages are sanitized and not leaked to `StatusText`.

## Code Example

```asp
<%
Dim oHTTP
Set oHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
oHTTP.Open "POST", "https://example.com/api/submit", False
oHTTP.SetRequestHeader "Content-Type", "application/json"
oHTTP.Send "{\"key\":\"value\"}"
If oHTTP.Status = 200 Then
    Response.Write oHTTP.ResponseText
Else
    Response.Write "Error: " & oHTTP.Status & " " & oHTTP.StatusText
End If
Set oHTTP = Nothing
%>
```