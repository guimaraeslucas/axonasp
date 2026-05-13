# GetComponent

## Overview

Returns a `G3ALComponentProxy` native object, providing granular manipulation methods to modify a specific DOM element's classes, styles, and attributes without having to re-render its entire HTML.

## Syntax

```vbscript
Set objProxy = objAxonLive.GetComponent(componentId)
```

```javascript
var objProxy = objAxonLive.GetComponent(componentId);
```

## Parameters and Arguments

* **componentId** (`String`): The HTML ID of the component element to proxy.

## Return Values

Returns an `Object` (`G3ALComponentProxy`).

## Remarks

The returned proxy object acts as a bridge to manipulate the DOM element on the client side. The following methods are available on the proxy object:
* `SetStyle(attributeName, attributeValue)`
* `AddClass(className)`
* `RemoveClass(className)`
* `SetAttribute(attributeName, attributeValue)`
* `RemoveAttribute(attributeName)`
* `AddTitle(titleText)`
* `RemoveTitle()`
* `SetValue(value)`

## Code Example

### VBScript
```vbscript
Dim AxonLive, btnAction
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

If AxonLive.IsAsyncRequest Then
    Set btnAction = AxonLive.GetComponent("btnAction")
    btnAction.AddClass "btn-success"
    btnAction.SetAttribute "disabled", "disabled"
    btnAction.SetValue "Action Completed"
    
    AxonLive.EndAsyncResponse()
End If
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

if (AxonLive.IsAsyncRequest) {
    var btnAction = AxonLive.GetComponent("btnAction");
    btnAction.AddClass("btn-success");
    btnAction.SetAttribute("disabled", "disabled");
    btnAction.SetValue("Action Completed");
    
    AxonLive.EndAsyncResponse();
}
```