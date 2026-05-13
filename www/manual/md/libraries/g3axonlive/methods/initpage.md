# InitPage

## Overview

Initializes the page for the AxonLive framework. It parses the incoming request to determine if it is an async POST event and automatically registers the current page URL for session tracking.

## Syntax

```vbscript
objAxonLive.InitPage()
```

```javascript
objAxonLive.InitPage();
```

## Parameters and Arguments

None.

## Return Values

Returns a `Boolean`. Returns `True` if it successfully initialized an asynchronous event, or `False` if it initialized a standard page load.

## Remarks

You MUST call this method at the very top of your ASP script, before checking `IsAsyncRequest` or attempting to read any event properties. Calling it multiple times in the same script has no adverse effects.

## Code Example

### VBScript
```vbscript
Dim AxonLive
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage() ' Mandatory initialization
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage(); // Mandatory initialization
```