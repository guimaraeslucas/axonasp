# EventArgs

## Overview

Returns the entire event arguments map sent by the client, encoded as a JSON string. This is useful for passing complex contextual data from the frontend to the backend during an asynchronous event.

## Syntax

```vbscript
strJsonArgs = objAxonLive.EventArgs
```

```javascript
var strJsonArgs = objAxonLive.EventArgs;
```

## Return Values

Returns a `String` containing a JSON object representation of the arguments. Returns `{}` if there are no arguments.

## Remarks

The client bridge (`g3axonlive.js`) automatically collects `data-g3al-arg-*` attributes from the HTML element that triggered the event. For example, `data-g3al-arg-step="5"` will be passed as `{"step": "5"}` in the JSON payload.

## Code Example

### VBScript
```vbscript
Dim AxonLive, jsonArgs
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

If AxonLive.IsAsyncRequest Then
    jsonArgs = AxonLive.EventArgs
    ' You can parse jsonArgs using the G3JSON library
    Response.Write "Received JSON: " & jsonArgs
    AxonLive.EndAsyncResponse()
End If
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

if (AxonLive.IsAsyncRequest) {
    var jsonArgs = AxonLive.EventArgs;
    Response.Write("Received JSON: " + jsonArgs);
    AxonLive.EndAsyncResponse();
}
```