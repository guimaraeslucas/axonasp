# ServerXMLHTTP.SetTimeouts Method

Configures the timeout values (in milliseconds) for DNS resolution, server connection, request sending, and response receiving.

## Syntax

```asp
objHTTP.SetTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
```

## Parameters

| Parameter | Type | Description |
|---|---|---|
| `resolveTimeout` | Integer | Maximum milliseconds for name resolution (`0` for infinite). |
| `connectTimeout` | Integer | Maximum milliseconds for establishing a connection (`0` for infinite). |
| `sendTimeout` | Integer | Maximum milliseconds for sending the request (`0` for infinite). |
| `receiveTimeout` | Integer | Maximum milliseconds for receiving the response (`0` for infinite). |

## Return Value

Empty. This method does not return a value.

## Remarks

- Call `SetTimeouts` before `Send` (typically before `Open`).
- Passing any negative value raises `E_INVALIDARG` (`0x80070057`).
- When an operation exceeds its configured timeout, `Send` raises `0x80072EE2` (`-2147012894`, `"The operation timed out"`).
- Method names are case-insensitive.

## Code Example

```asp
<%
On Error Resume Next
Dim oHTTP
Set oHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
oHTTP.SetTimeouts 1000, 1000, 1000, 1000
oHTTP.Open "GET", "https://example.com/slow-endpoint", False
oHTTP.Send
If Err.Number = -2147012894 Then
    Response.Write "Request timed out."
ElseIf Err.Number <> 0 Then
    Response.Write "Connection failed: " & Err.Description
End If
Set oHTTP = Nothing
%>
```
