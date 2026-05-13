# GetComponentState

## Overview

Returns a diagnostic string listing all stored properties for a specific component and session.

## Syntax

```vbscript
strDump = objAxonLive.GetComponentState(sessionID, componentID)
```

```javascript
var strDump = objAxonLive.GetComponentState(sessionID, componentID);
```

## Parameters and Arguments

* **sessionID** (`String`): The active Session ID.
* **componentID** (`String`): The component identifier.

## Return Values

Returns a `String` containing a formatted list of all property keys, their values, and their last updated timestamps.

## Remarks

This method is intended primarily for debugging purposes. Do not rely on its specific string format for parsing business logic.

## Code Example

### VBScript
```vbscript
Dim AxonLive, sID, dump
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

sID = Session.SessionID

dump = AxonLive.GetComponentState(sID, "wizard")
Response.Write "<pre>" & Server.HTMLEncode(dump) & "</pre>"
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

var sID = Session.SessionID;

var dump = AxonLive.GetComponentState(sID, "wizard");
Response.Write("<pre>" + Server.HTMLEncode(dump) + "</pre>");
```