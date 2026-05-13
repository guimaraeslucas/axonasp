# GetComponentProperty

## Overview

Retrieves a property value from the persistent global state for a given session and component.

## Syntax

```vbscript
strValue = objAxonLive.GetComponentProperty(sessionID, componentID, propertyName)
```

```javascript
var strValue = objAxonLive.GetComponentProperty(sessionID, componentID, propertyName);
```

## Parameters and Arguments

* **sessionID** (`String`): The active Session ID.
* **componentID** (`String`): The component identifier used when saving the property.
* **propertyName** (`String`): The name of the property to retrieve.

## Return Values

Returns a `String` containing the stored value. Returns `Empty` (VBScript) or `undefined` (JavaScript) if the property does not exist.

## Remarks

This method retrieves data from the optimized Go memory map, bypassing standard `Session` state. It's the primary way to re-hydrate your UI during a standard page load or to check state during an async event.

## Code Example

### VBScript
```vbscript
Dim AxonLive, sID, stepVal
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

sID = Session.SessionID

stepVal = AxonLive.GetComponentProperty(sID, "wizard", "step")
If IsEmpty(stepVal) Then stepVal = "1"

Response.Write "Current Step: " & stepVal
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

var sID = Session.SessionID;

var stepVal = AxonLive.GetComponentProperty(sID, "wizard", "step");
if (!stepVal) stepVal = "1";

Response.Write("Current Step: " + stepVal);
```