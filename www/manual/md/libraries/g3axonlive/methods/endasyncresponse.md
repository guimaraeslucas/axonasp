# EndAsyncResponse

## Overview

Finalizes the asynchronous event cycle. It collects all registered HTML component patches and scheduled client actions, serializes them into a JSON payload, writes the response to the client, and safely halts the execution of the ASP script.

## Syntax

```vbscript
objAxonLive.EndAsyncResponse()
```

```javascript
objAxonLive.EndAsyncResponse();
```

## Parameters and Arguments

None.

## Return Values

Returns `Empty`. 

## Remarks

Calling this method is mandatory at the end of your `If AxonLive.IsAsyncRequest` block. Once called, the ASP engine immediately stops processing the file (similar to `Response.End`), ensuring that the background request only receives the JSON instructions and not the remaining HTML template of your page.

## Code Example

### VBScript
```vbscript
Dim AxonLive
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

If AxonLive.IsAsyncRequest Then
    ' 1. Process Logic
    ' 2. Register Component Patches
    AxonLive.RegisterComponent "myDiv", "<div>Updated HTML</div>"
    
    ' 3. Finish and Halt
    AxonLive.EndAsyncResponse()
End If
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

if (AxonLive.IsAsyncRequest) {
    // 1. Process Logic
    // 2. Register Component Patches
    AxonLive.RegisterComponent("myDiv", "<div>Updated HTML</div>");
    
    // 3. Finish and Halt
    AxonLive.EndAsyncResponse();
}
```