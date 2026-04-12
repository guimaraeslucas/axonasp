# ASP Dictionary Object

## Overview

ASP applications commonly use Scripting.Dictionary as a key/value container for in-memory data structures, JSON-like object maps, and lookup tables.

## Syntax

```asp
Set dict = Server.CreateObject("Scripting.Dictionary")
```

## Methods

- Add(key, value)
- Exists(key)
- Remove(key)
- RemoveAll()
- Keys()
- Items()
- Item(key)
- Key(oldKey) = newKey
- Count

## Properties

- Count (read-only)
- CompareMode (read/write)

## Code Example

```asp
<%
Option Explicit
Dim dict
Set dict = Server.CreateObject("Scripting.Dictionary")
dict.CompareMode = 1
dict.Add "name", "AxonASP"
dict.Add "version", "2"
If dict.Exists("name") Then
    Response.Write dict.Item("name")
End If
Set dict = Nothing
%>
```
