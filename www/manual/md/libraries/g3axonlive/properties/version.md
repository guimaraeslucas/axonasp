# Version

## Overview

Returns the current version of the AxonLive framework running on the server.

## Syntax

```vbscript
strVersion = objAxonLive.Version
```

```javascript
var strVersion = objAxonLive.Version;
```

## Return Values

Returns a `String` containing the version number (e.g., `"2.0.0"`).

## Remarks

Useful for debugging or ensuring compatibility with specific client-side `g3axonlive.js` scripts.

## Code Example

### VBScript
```vbscript
Dim AxonLive
Set AxonLive = Server.CreateObject("G3AXONLIVE")
Response.Write "AxonLive Version: " & AxonLive.Version
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");
Response.Write("AxonLive Version: " + AxonLive.Version);
```