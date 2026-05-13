# EventComponentID

## Overview

Returns the ID of the component that fired the current asynchronous event.

## Syntax

```vbscript
strComponentID = objAxonLive.EventComponentID
```

```javascript
var strComponentID = objAxonLive.EventComponentID;
```

## Return Values

Returns a `String` representing the ID of the component. If the request is not an asynchronous event, it returns an empty string.

## Remarks

This property is crucial for routing logic inside the `If AxonLive.IsAsyncRequest` block. You use it in a `Select Case` or `switch` statement to determine which piece of logic should execute.

## Code Example

### VBScript
```vbscript
Dim AxonLive, compID
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

If AxonLive.IsAsyncRequest Then
    compID = AxonLive.EventComponentID
    If compID = "btnIncrement" Then
        ' Handle increment logic
    End If
    AxonLive.EndAsyncResponse()
End If
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

if (AxonLive.IsAsyncRequest) {
    var compID = AxonLive.EventComponentID;
    switch (compID) {
        case "btnIncrement":
            // Handle increment logic
            break;
    }
    AxonLive.EndAsyncResponse();
}
```