# Use ASP Cookies

## Overview
Cookies in G3Pix AxonASP are used to store small amounts of data on the client's browser. They are essential for maintaining user state, persisting preferences, and identifying sessions across multiple requests. In AxonASP, cookies are managed through the **Request.Cookies** collection (for reading) and the **Response.Cookies** collection (for writing).

## Prerequisites
- **Client Support**: The client browser must be configured to accept cookies.
- **Header Synchronization**: Cookie modifications via `Response.Cookies` must occur before any HTML or data is flushed to the client, unless buffering is enabled.

## How it Works
Cookies are passed between the server and the client via HTTP headers.
- **Reading**: When a browser makes a request, it includes all cookies associated with the domain. AxonASP populates the **Request.Cookies** collection with these values.
- **Writing**: When you set a value in **Response.Cookies**, AxonASP includes a `Set-Cookie` header in the HTTP response.
- **Subkeys**: A single cookie can store multiple key-value pairs, known as subkeys, which helps organize data without exceeding browser cookie limits.

## API Reference

### Request.Cookies
- **Item(name)**: Returns a **String** containing the cookie value.
- **HasKeys**: Returns a **Boolean** indicating if the cookie contains subkeys.

### Response.Cookies
- **Domain**: Sets a **String** specifying the domain for which the cookie is valid.
- **Expires**: Sets a **Date** specifying the expiration date and time of the cookie.
- **HttpOnly**: Sets a **Boolean** to prevent client-side script access to the cookie (Security feature).
- **Path**: Sets a **String** specifying the subset of paths on the server for which the cookie is valid.
- **Secure**: Sets a **Boolean** indicating if the cookie should only be transmitted over HTTPS.

## Code Example
The following example demonstrates how to write a persistent cookie with subkeys and then read one of the values back.

```asp
<%
' Writing a cookie with subkeys and an expiration date
Response.Cookies("UserSettings")("Theme") = "Dark"
Response.Cookies("UserSettings")("FontSize") = "14px"
Response.Cookies("UserSettings").Expires = DateAdd("m", 1, Now())
Response.Cookies("UserSettings").HttpOnly = True

' Reading a cookie value
Dim currentTheme
currentTheme = Request.Cookies("UserSettings")("Theme")

If currentTheme <> "" Then
    Response.Write "Current User Theme: " & currentTheme
End If
%>
```

## Remarks
- **Size Limits**: Browsers typically limit cookies to 4KB each. Use them for identifiers or small settings, not large datasets.
- **Security**: Always set the **HttpOnly** property to **True** for sensitive cookies to mitigate cross-site scripting (XSS) risks.
- **Expiration**: If no expiration date is set, the cookie is treated as a session cookie and will be deleted when the browser is closed.
