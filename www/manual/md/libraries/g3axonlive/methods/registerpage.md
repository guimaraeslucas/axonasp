# RegisterPage

## Overview

Records the ASP script URL for a session so the backend framework knows which page to re-execute when an asynchronous event arrives from the client.

## Syntax

```vbscript
objAxonLive.RegisterPage sessionID, scriptURL
```

```javascript
objAxonLive.RegisterPage(sessionID, scriptURL);
```

## Parameters and Arguments

* **sessionID** (`String`): The active Session ID.
* **scriptURL** (`String`): The URL of the script (usually `Request.ServerVariables("SCRIPT_NAME")`).

## Return Values

Returns `Empty`.

## Remarks

You generally **do not need to call this manually**. The `InitPage` method automatically registers the current page on your behalf during standard page loads and async events. This method is exposed for edge cases where you might need to route async requests to a different handler script explicitly.

## Code Example

### VBScript
```vbscript
Dim AxonLive, sID
Set AxonLive = Server.CreateObject("G3AXONLIVE")

sID = Session.SessionID
AxonLive.RegisterPage sID, "/custom_handler.asp"
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");

var sID = Session.SessionID;
AxonLive.RegisterPage(sID, "/custom_handler.asp");
```