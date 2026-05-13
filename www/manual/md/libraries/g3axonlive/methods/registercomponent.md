# RegisterComponent

## Overview

Registers an updated HTML block for a specific component. When `EndAsyncResponse` is called, the framework sends this HTML patch to the browser, replacing the component's `outerHTML`.

## Syntax

```vbscript
objAxonLive.RegisterComponent componentId, htmlString
```

```javascript
objAxonLive.RegisterComponent(componentId, htmlString);
```

## Parameters and Arguments

* **componentId** (`String`): The ID of the HTML element on the client side to be replaced.
* **htmlString** (`String`): The completely new HTML string that will replace the component. Note: this replaces the `outerHTML`, so the root tag should typically maintain the same ID.

## Return Values

Returns `Empty`.

## Remarks

The `htmlString` must be a valid HTML string. If you replace an element, ensure the new HTML includes the same `id` and `data-g3al-*` attributes if you want it to remain interactive for future events.

## Code Example

### VBScript
```vbscript
Dim AxonLive, newHtml
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

If AxonLive.IsAsyncRequest Then
    newHtml = "<div id=""statusMessage"" class=""alert alert-success"">Saved!</div>"
    AxonLive.RegisterComponent "statusMessage", newHtml
    AxonLive.EndAsyncResponse()
End If
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

if (AxonLive.IsAsyncRequest) {
    var newHtml = "<div id=\"statusMessage\" class=\"alert alert-success\">Saved!</div>";
    AxonLive.RegisterComponent("statusMessage", newHtml);
    AxonLive.EndAsyncResponse();
}
```