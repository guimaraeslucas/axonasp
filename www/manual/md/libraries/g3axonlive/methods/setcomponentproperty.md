# SetComponentProperty

## Overview

Stores a property value in the persistent global state for a given session and component. This data persists across all asynchronous calls and page reloads for the user.

## Syntax

```vbscript
objAxonLive.SetComponentProperty sessionID, componentID, propertyName, propertyValue
```

```javascript
objAxonLive.SetComponentProperty(sessionID, componentID, propertyName, propertyValue);
```

## Parameters and Arguments

* **sessionID** (`String`): The active Session ID (usually obtained via `Session.SessionID`).
* **componentID** (`String`): An arbitrary component identifier acting as a namespace bucket.
* **propertyName** (`String`): The name of the property to store.
* **propertyValue** (`String`): The value to store.

## Return Values

Returns `Empty`.

## Remarks

Unlike standard ASP `Session` variables, this state is maintained in an optimized Go `sync.RWMutex` map, offering excellent concurrent read/write performance.

## Code Example

### VBScript
```vbscript
Dim AxonLive, sID
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

sID = Session.SessionID

If AxonLive.IsAsyncRequest Then
    ' Save the user's progress
    AxonLive.SetComponentProperty sID, "wizard", "step", "2"
    AxonLive.EndAsyncResponse()
End If
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

var sID = Session.SessionID;

if (AxonLive.IsAsyncRequest) {
    // Save the user's progress
    AxonLive.SetComponentProperty(sID, "wizard", "step", "2");
    AxonLive.EndAsyncResponse();
}
```