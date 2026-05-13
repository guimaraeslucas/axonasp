# RemoveComponentProperty

## Overview

Deletes a specific property entry from the persistent global state.

## Syntax

```vbscript
objAxonLive.RemoveComponentProperty sessionID, componentID, propertyName
```

```javascript
objAxonLive.RemoveComponentProperty(sessionID, componentID, propertyName);
```

## Parameters and Arguments

* **sessionID** (`String`): The active Session ID.
* **componentID** (`String`): The component identifier.
* **propertyName** (`String`): The name of the property to remove.

## Return Values

Returns `Empty`.

## Remarks

Use this method to clean up individual properties without affecting the rest of the component's state.

## Code Example

### VBScript
```vbscript
Dim AxonLive, sID
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

sID = Session.SessionID

If AxonLive.IsAsyncRequest Then
    If AxonLive.EventComponentID = "btnResetStep" Then
        AxonLive.RemoveComponentProperty sID, "wizard", "step"
    End If
    AxonLive.EndAsyncResponse()
End If
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXON.LIVE");
AxonLive.InitPage();

var sID = Session.SessionID;

if (AxonLive.IsAsyncRequest) {
    if (AxonLive.EventComponentID === "btnResetStep") {
        AxonLive.RemoveComponentProperty(sID, "wizard", "step");
    }
    AxonLive.EndAsyncResponse();
}
```