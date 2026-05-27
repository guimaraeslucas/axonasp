# G3FILEUPLOADER Properties

## Overview
This page summarizes properties exposed by **G3FILEUPLOADER** in G3Pix AxonASP.

## Properties Reference

| Property | Access | Type | Description |
|---|---|---|---|
| AllowAbsolutePaths | Read/Write | Boolean | When true, allows saving files to absolute system paths outside the web root. |
| AllowedExtensions | Read-only | Array of String | Gets the current allowed extension set. |
| BlockedExtensions | Read-only | Array of String | Gets the current blocked extension set. |
| DebugMode | Read/Write | Boolean | Gets or sets whether debug information is logged. |
| FormFields | Read-only | Dictionary | Returns a Dictionary containing all non-file form fields sent in the request. |
| MaxFileSize | Read/Write | Integer | Gets or sets the maximum upload size in bytes for each file. |
| PreserveOriginalName | Read/Write | Boolean | Gets or sets whether saved files preserve the original uploaded filename. |

## Remarks
- Property names are case-insensitive.
- Setting `AllowAbsolutePaths` to **True** should be done with caution as it bypasses the sandbox restriction.
