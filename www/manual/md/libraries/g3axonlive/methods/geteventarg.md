# GetEventArg

## Overview

Retrieves the value of a specific event argument sent by the client during the async request.

## Syntax

```vbscript
strArgValue = objAxonLive.GetEventArg(argName)
```

```javascript
var strArgValue = objAxonLive.GetEventArg(argName);
```

## Parameters and Arguments

* **argName** (`String`): The name of the argument to retrieve.

## Return Values

Returns a `String` containing the argument value. If the argument does not exist, it returns `Empty` (VBScript) or `undefined`/empty string (JavaScript).

## Remarks

Arguments are automatically populated by the `g3axonlive.js` bridge from elements containing `data-g3al-arg-[name]="[value]"` attributes.

## Code Example

### VBScript
```vbscript
Dim AxonLive, stepVal
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

If AxonLive.IsAsyncRequest Then
    ' Assuming client element had: data-g3al-arg-step="5"
    stepVal = AxonLive.GetEventArg("step")
    If stepVal = "" Then stepVal = "1"
    
    Response.Write "Moving " & stepVal & " steps."
    AxonLive.EndAsyncResponse()
End If
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

if (AxonLive.IsAsyncRequest) {
    // Assuming client element had: data-g3al-arg-step="5"
    var stepVal = AxonLive.GetEventArg("step");
    if (!stepVal) stepVal = "1";
    
    Response.Write("Moving " + stepVal + " steps.");
    AxonLive.EndAsyncResponse();
}
```