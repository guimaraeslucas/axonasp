# AxPoweredByImage

## Overview

Use `AxPoweredByImage` to read the built-in "Powered by AxonASP" promotional image as a Base64 data URI string.

## Prerequisites

- Instantiate the library with `Server.CreateObject("G3AXON.FUNCTIONS")`.

## Syntax

```asp
imageUri = obj.AxPoweredByImage()
```

## Parameters

- This method does not require parameters.

## Return Value

- **String**: Returns a Base64 encoded data URI representing a PNG image (e.g., `data:image/png;base64,...`).

## Remarks

- Use this method to easily embed a "Powered by AxonASP" badge into your HTML output without relying on external image files.
- The returned string can be used directly as the `src` attribute of an HTML `<img>` tag.
- Method names are case-insensitive in VBScript dispatch.

## Example

```asp
<%
Option Explicit
Dim ax, imageUri

Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

imageUri = ax.AxPoweredByImage()
Response.Write "<img src=""" & imageUri & """ alt=""Powered by AxonASP"" />"

Set ax = Nothing
%>
```

## API Reference

- **Object**: `G3AXON.FUNCTIONS`
- **Method**: `AxPoweredByImage`
- **Arguments**: none
- **Returns**: `String` (Base64 image data URI)
