# SetTimer

## Overview

Queues a client action that instructs the browser to trigger a specific event on a component after a set delay. This is useful for building polling mechanisms, delayed redirects, or delayed UI resets.

## Syntax

```vbscript
objAxonLive.SetTimer componentId, eventName, delayMs
```

```javascript
objAxonLive.SetTimer(componentId, eventName, delayMs);
```

## Parameters and Arguments

* **componentId** (`String`): The ID of the component that will receive the delayed event.
* **eventName** (`String`): The name of the event to trigger (e.g., `ontimer`).
* **delayMs** (`Integer`): The delay in milliseconds before the event fires.

## Return Values

Returns `Empty`.

## Remarks

The client browser handles the actual timer via `setTimeout`. Once the time elapses, `g3axonlive.js` dispatches an asynchronous request back to the server just as if the user had manually interacted with the component.

## Code Example

### VBScript
```vbscript
Dim AxonLive
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

If AxonLive.IsAsyncRequest Then
    If AxonLive.EventComponentID = "btnStart" Then
        ' Trigger the 'ontimer' event on the 'timer1' component after 5 seconds
        AxonLive.SetTimer "timer1", "ontimer", 5000
    End If
    AxonLive.EndAsyncResponse()
End If
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

if (AxonLive.IsAsyncRequest) {
    if (AxonLive.EventComponentID === "btnStart") {
        // Trigger the 'ontimer' event on the 'timer1' component after 5 seconds
        AxonLive.SetTimer("timer1", "ontimer", 5000);
    }
    AxonLive.EndAsyncResponse();
}
```