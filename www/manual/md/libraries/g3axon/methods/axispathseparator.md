# Axispathseparator Method

## Overview

Checks whether a single-character input is a valid path separator on the current platform.

## Syntax

```asp
result = obj.Axispathseparator(character)
```

## Parameters

- character (String): Single character to validate.

## Return Values

- Boolean: True when the character is a path separator; otherwise False.

## Code Example

```asp
<%
Option Explicit
Dim obj
Set obj = Server.CreateObject("G3AXON.Functions")
Response.Write CStr(obj.Axispathseparator("/")) & "<br>"
Response.Write CStr(obj.Axispathseparator("a"))
Set obj = Nothing
%>
```