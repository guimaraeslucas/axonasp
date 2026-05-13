# IsAsyncRequest

## Overview

A boolean property that indicates whether the current page load is a standard full-page request or a background JSON `fetch` request triggered by an AxonLive component.

## Syntax

```vbscript
bIsAsync = objAxonLive.IsAsyncRequest
```

```javascript
var bIsAsync = objAxonLive.IsAsyncRequest;
```

## Return Values

Returns a `Boolean`. `True` if the request is an asynchronous event POST, otherwise `False`.

## Remarks

This property is the main gatekeeper for AxonLive logic. All component-handling logic must be placed inside a conditional block checking this property to avoid running during a standard page render.

## Code Example

### VBScript
```vbscript
Dim AxonLive
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

If AxonLive.IsAsyncRequest Then
    ' We are handling an AJAX background event.
    ' Process logic and stop execution.
    AxonLive.EndAsyncResponse()
End If

' If we get here, it's a normal page load. Render HTML below.
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

if (AxonLive.IsAsyncRequest) {
    // We are handling an AJAX background event.
    // Process logic and stop execution.
    AxonLive.EndAsyncResponse();
}

// If we get here, it's a normal page load. Render HTML below.
```