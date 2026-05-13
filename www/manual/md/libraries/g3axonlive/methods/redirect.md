# Redirect

## Overview

Queues a client action that securely navigates the browser to the specified URL. Since the request is asynchronous, a standard `Response.Redirect` will not work. This method instructs the frontend bridge to change `window.location.href`.

## Syntax

```vbscript
objAxonLive.Redirect url
```

```javascript
objAxonLive.Redirect(url);
```

## Parameters and Arguments

* **url** (`String`): The destination URL. It can be absolute or relative.

## Return Values

Returns `Empty`.

## Remarks

This method should be used instead of `Response.Redirect` when responding to an AxonLive event. The redirect happens immediately after the client bridge processes the JSON response.

## Code Example

### VBScript
```vbscript
Dim AxonLive
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

If AxonLive.IsAsyncRequest Then
    If AxonLive.EventComponentID = "btnLogout" Then
        ' Perform logout logic here, then redirect
        AxonLive.Redirect "/login.asp"
    End If
    AxonLive.EndAsyncResponse()
End If
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

if (AxonLive.IsAsyncRequest) {
    if (AxonLive.EventComponentID === "btnLogout") {
        // Perform logout logic here, then redirect
        AxonLive.Redirect("/login.asp");
    }
    AxonLive.EndAsyncResponse();
}
```