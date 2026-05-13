# EventName

## Overview

Returns the name of the asynchronous event that was fired by the component (e.g., `onclick`, `onchange`).

## Syntax

```vbscript
strEventName = objAxonLive.EventName
```

```javascript
var strEventName = objAxonLive.EventName;
```

## Return Values

Returns a `String` representing the name of the event. Returns an empty string if the request is not asynchronous.

## Remarks

While `EventComponentID` tells you *who* triggered the request, `EventName` tells you *what* they did. This is useful when a single component listens to multiple event types (e.g., handling both `onclick` and `ondblclick`).

## Code Example

### VBScript
```vbscript
Dim AxonLive, evtName
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

If AxonLive.IsAsyncRequest Then
    evtName = AxonLive.EventName
    If evtName = "onclick" Then
        ' Handle click event
    End If
    AxonLive.EndAsyncResponse()
End If
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

if (AxonLive.IsAsyncRequest) {
    var evtName = AxonLive.EventName;
    if (evtName === "onclick") {
        // Handle click event
    }
    AxonLive.EndAsyncResponse();
}
```