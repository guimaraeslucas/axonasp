# Retrieve Default Configured Logo

## Overview

Retrieves the server's configured default logo as an inline Base64 data URI string.

## Syntax

```vbscript
strDataUri = obj.axgetlogo()
```

## Parameters

- None.

## Return Value

String. The data URI containing the encoded logo image.

## Remarks

The logo source file is defined in the `axfunctions.ax_default_logo_path` property within the AxonASP configuration file. Returns an empty string if the file is missing or invalid.

## Code Example

```vbscript
Dim obj, strUri
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")
strUri = obj.axgetlogo()
Response.Write "<img src=""" & strUri & """>"
```