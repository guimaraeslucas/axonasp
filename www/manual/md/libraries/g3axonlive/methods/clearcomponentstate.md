# ClearComponentState

## Overview

Clears all properties saved in the persistent store for a given session and component pair.

## Syntax

```vbscript
objAxonLive.ClearComponentState sessionID, componentID
```

```javascript
objAxonLive.ClearComponentState(sessionID, componentID);
```

## Parameters and Arguments

* **sessionID** (`String`): The active Session ID.
* **componentID** (`String`): The component identifier.

## Return Values

Returns `Empty`.

## Remarks

This method deletes all keys sharing the same `sessionID` and `componentID` prefix in the internal memory map. It is useful for resetting a complex component (like a wizard or a multi-field form) to its factory state.

## Code Example

### VBScript
```vbscript
Dim AxonLive, sID
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

sID = Session.SessionID

If AxonLive.IsAsyncRequest Then
    If AxonLive.EventComponentID = "btnResetAll" Then
        AxonLive.ClearComponentState sID, "wizard"
    End If
    AxonLive.EndAsyncResponse()
End If
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

var sID = Session.SessionID;

if (AxonLive.IsAsyncRequest) {
    if (AxonLive.EventComponentID === "btnResetAll") {
        AxonLive.ClearComponentState(sID, "wizard");
    }
    AxonLive.EndAsyncResponse();
}
```