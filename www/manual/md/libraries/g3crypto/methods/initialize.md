# Initialize Method

## Overview

Resets the internal state of the G3Pix AxonASP G3CRYPTO object, clearing any cached hash results and configuration.

## Syntax

```asp
obj.Initialize()
```

## Parameters

This method accepts no parameters.

## Return Values

Returns an Empty value (VT_EMPTY).

## Remarks

- Instantiated via `Server.CreateObject("G3CRYPTO")`.
- Calling this method clears the `Hash` and `HashSize` properties by removing the reference to the last computed hash.
- This is useful for ensuring no sensitive data remains in memory between different cryptographic operations using the same object instance.

## Code Example

```asp
<%
Dim crypto
Set crypto = Server.CreateObject("G3CRYPTO")
crypto.ComputeHash("Some data", "sha256")
' Clear memory and internal state
crypto.Initialize()
Set crypto = Nothing
%>
```
