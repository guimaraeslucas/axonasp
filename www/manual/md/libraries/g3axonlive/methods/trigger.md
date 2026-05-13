# Trigger

## Overview

Queues a client action that immediately fires a specific client-side event without requiring user interaction.

## Syntax

```vbscript
objAxonLive.Trigger componentId, eventName
```

```javascript
objAxonLive.Trigger(componentId, eventName);
```

## Parameters and Arguments

* **componentId** (`String`): The ID of the component on which to trigger the event.
* **eventName** (`String`): The name of the event to fire (e.g., `onclick`, `onchange`).

## Return Values

Returns `Empty`.

## Remarks

This is useful for chaining server-side workflows. For instance, successfully saving a form could automatically trigger a refresh event on a separate grid component.

## Code Example

### VBScript
```vbscript
Dim AxonLive
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

If AxonLive.IsAsyncRequest Then
    If AxonLive.EventComponentID = "btnSave" Then
        ' Save logic here...
        
        ' Now tell the client to automatically "click" the refresh button
        AxonLive.Trigger "btnRefreshData", "onclick"
    End If
    AxonLive.EndAsyncResponse()
End If
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

if (AxonLive.IsAsyncRequest) {
    if (AxonLive.EventComponentID === "btnSave") {
        // Save logic here...
        
        // Now tell the client to automatically "click" the refresh button
        AxonLive.Trigger("btnRefreshData", "onclick");
    }
    AxonLive.EndAsyncResponse();
}
```