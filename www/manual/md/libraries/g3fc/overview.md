# Use the G3FC Archive Library in G3Pix AxonASP

## Overview
Use the G3FC library to create, inspect, search, and extract G3FC archives in G3Pix AxonASP. This page provides a concise API summary for object creation, available methods, and property behavior.

## Syntax
```asp
Set fc = Server.CreateObject("G3FC")
```

## Parameters and Arguments
- ProgID (String, Required): `G3FC`.
- Object variable (Object, Required): Receives the instantiated native G3FC object handle.
- Method arguments (Depends on method): Each method validates argument count and argument types at runtime.

## Return Values
- `Server.CreateObject("G3FC")` returns a native object handle bound to the G3FC implementation.
- If object creation fails, the call returns `Empty` and `Server.GetLastError` contains the failure details.

## Remarks
- **Object creation is case-insensitive for supported ProgIDs.** `G3FC`, `g3fc`, and mixed-case forms resolve to the same native implementation.
- **Unsupported ProgIDs fall through to host-level CreateObject handling.** If the host cannot create the object, AxonASP raises `AxonASP cannot create object`.
- **Compatibility aliases are mapped to native AxonASP implementations where available.** For G3FC specifically, use the primary ProgID `G3FC`.
- **Method names are case-insensitive.**
- G3FC exposes no dedicated properties. Property-style access is routed to method dispatch.

## Code Example
```asp
<%
Option Explicit

Dim fc
Set fc = Server.CreateObject("G3FC")

If IsObject(fc) Then
	Response.Write "G3FC object created successfully."
Else
	Response.Write "Failed to create G3FC object."
End If

Set fc = Nothing
%>
```

## Methods Reference

| Method | Returns | Description |
|---|---|---|
| Create | Boolean | Returns `True` when the archive is created successfully; returns `False` when required arguments are missing, path resolution fails, or archive creation fails. |
| Extract | Boolean | Returns `True` when all archive entries are extracted to the target directory; returns `False` when arguments are invalid or extraction fails. |
| List | Array of `Scripting.Dictionary` or Empty | Returns an Array where each item is a `Scripting.Dictionary` describing one archive entry (for example, path and size metadata). Returns `Empty` when arguments are invalid or archive reading fails. |
| Info | Boolean | Returns `True` when metadata export to the output file completes; returns `False` when arguments are invalid or metadata export fails. |
| Find | Array of `Scripting.Dictionary` or Empty | Returns an Array where each item is a `Scripting.Dictionary` for an entry matched by substring or regular expression. Returns `Empty` when arguments are invalid or search/read fails. |
| ExtractSingle | Boolean | Returns `True` when the requested single entry is extracted to the target path; returns `False` when arguments are invalid, the entry does not exist, or extraction fails. |

## Properties Reference

| Property | Access | Type | Description |
|---|---|---|---|
| None | Not applicable | Not applicable | G3FC does not expose dedicated properties. Use method calls for all operations. |

## API Reference
- Object creation: `Server.CreateObject("G3FC")`
- Method aliases: `ExtractSingle` also accepts `extract-single` and `extract_single`
- Member resolution: case-insensitive dispatch for methods and property-style calls

