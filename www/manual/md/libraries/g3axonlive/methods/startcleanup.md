# StartCleanup

## Overview

Starts a background process (goroutine) in the Go engine to clean up idle AxonLive session data.

## Syntax

```vbscript
objAxonLive.StartCleanup()
```

```javascript
objAxonLive.StartCleanup();
```

## Parameters and Arguments

None.

## Return Values

Returns `Empty`.

## Remarks

This method is usually triggered automatically during server startup if `g3axonlive_active` is enabled in `axonasp.toml`. You generally do not need to call this manually unless you previously called `StopCleanup`.

## Code Example

### VBScript
```vbscript
Dim AxonLive
Set AxonLive = Server.CreateObject("G3AXONLIVE")

' Explicitly start the memory cleanup process
AxonLive.StartCleanup()
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");

// Explicitly start the memory cleanup process
AxonLive.StartCleanup();
```