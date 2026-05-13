# StopCleanup

## Overview

Stops the background process that cleans up idle AxonLive session data.

## Syntax

```vbscript
objAxonLive.StopCleanup()
```

```javascript
objAxonLive.StopCleanup();
```

## Parameters and Arguments

None.

## Return Values

Returns `Empty`.

## Remarks

This method is typically only used during server shutdown or in specialized environments where you want to manage memory cleanup manually.

## Code Example

### VBScript
```vbscript
Dim AxonLive
Set AxonLive = Server.CreateObject("G3AXONLIVE")

' Explicitly stop the memory cleanup process
AxonLive.StopCleanup()
```

### JavaScript
```javascript
var AxonLive = Server.CreateObject("G3AXONLIVE");

// Explicitly stop the memory cleanup process
AxonLive.StopCleanup();
```