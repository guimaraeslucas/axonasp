# RemoveSession

## Overview

Completely deletes all AxonLive state data associated with a specific session ID, including component values, page registration, and last access time.

## Syntax

```vbscript
objAxonLive.RemoveSession sessionID
```

```javascript
objAxonLive.RemoveSession(sessionID);
```

## Parameters and Arguments

* **sessionID** (`String`): The Session ID to remove.

## Return Values

Returns `Empty`.

## Remarks

This method is typically called when a user logs out or their session is explicitly ended. It ensures that the Go memory map frees up resources associated with that user immediately, rather than waiting for the background cleanup process.

## Code Example

### VBScript
```vbscript
Dim AxonLive, sID
Set AxonLive = Server.CreateObject("G3AXONLIVE")
AxonLive.InitPage()

sID = Session.SessionID

If Request.QueryString("action") = "logout" Then
    AxonLive.RemoveSession sID
    Session.Abandon
    Response.Redirect "/login.asp"
End If
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();

var sID = Session.SessionID;

if (Request.QueryString("action") == "logout") {
    AxonLive.RemoveSession(sID);
    Session.Abandon();
    Response.Redirect("/login.asp");
}
```