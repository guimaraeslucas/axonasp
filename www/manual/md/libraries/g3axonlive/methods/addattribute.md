# AddAttribute

## Overview

Queues a client action that adds or updates an HTML attribute on a specific component element in the browser DOM.

## Syntax

```vbscript
objAxonLive.AddAttribute componentId, attributeName, attributeValue
```

```javascript
objAxonLive.AddAttribute(componentId, attributeName, attributeValue);
```

## Parameters and Arguments

* **componentId** (`String`): The ID of the component to update.
* **attributeName** (`String`): The name of the HTML attribute (e.g., `disabled`, `data-count`).
* **attributeValue** (`String`): The value to assign to the attribute.

## Return Values

Returns `Empty`.

## Remarks

If you need to make several granular changes to an element, consider using `GetComponent` to obtain a `G3ALComponentProxy` instead.

## Code Example

### VBScript
```vbscript
Dim AxonLive
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

If AxonLive.IsAsyncRequest Then
    ' Disable the submit button to prevent double-clicks
    AxonLive.AddAttribute "btnSubmit", "disabled", "disabled"
    AxonLive.EndAsyncResponse()
End If
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

if (AxonLive.IsAsyncRequest) {
    // Disable the submit button to prevent double-clicks
    AxonLive.AddAttribute("btnSubmit", "disabled", "disabled");
    AxonLive.EndAsyncResponse();
}
```