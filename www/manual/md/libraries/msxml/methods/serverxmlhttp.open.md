# ServerXMLHTTP.Open Method

Configures the HTTP method, URL, and optional credentials for a pending request. Must be called before `Send`.

## Syntax

```asp
objHTTP.Open method, url [, async [, user [, password]]]
```

## Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `method` | String | Yes | The HTTP verb in uppercase (e.g., `"GET"`, `"POST"`, `"PUT"`, `"DELETE"`). The value is automatically converted to uppercase. |
| `url` | String | Yes | The absolute URL of the resource to request. |
| `async` | Boolean | No | Accepted for compatibility. The implementation always behaves synchronously regardless of this value. |
| `user` | String | No | Username for HTTP Basic authentication. |
| `password` | String | No | Password for HTTP Basic authentication. |

## Return Value

Empty. This method does not return a value.

- Calling `Open` resets the object state: `ReadyState` is set to 1 and any previously stored response data is cleared.
- All previously set request headers are preserved between `Open` calls.
- Method names are case-insensitive.
- Passing an unrecognized scheme (other than `http` or `https`) or a relative URL raises error `0x80072EE6` ("The URL does not use a recognized protocol").

## Code Example

```asp
<%
Dim oHTTP
Set oHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
oHTTP.Open "GET", "https://example.com/api/items", False
oHTTP.SetRequestHeader "Accept", "application/xml"
oHTTP.Send
If oHTTP.Status = 200 Then
    Response.Write oHTTP.ResponseText
End If
Set oHTTP = Nothing
%>
```