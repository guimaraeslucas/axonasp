# ConvertTextEncoding Method

## Overview
Returns a string that has been converted from one text encoding to another.

## Syntax
```asp
convertedText = files.ConvertTextEncoding(text, fromEnc, toEnc)
```

## Parameters and Arguments
- **text** (String, Required): The source text to be converted.
- **fromEnc** (String, Required): The current encoding of the text.
- **toEnc** (String, Required): The target encoding for the result.

## Return Values
Returns a **String** containing the converted text.

## Remarks
- This method is performed in memory and does not affect any files on disk.
- It is useful for processing data received from external sources with different encoding requirements.

## Code Example
```asp
<%
Dim files, sourceText, targetText
Set files = Server.CreateObject("G3FILES")
sourceText = "Special characters: é à ö"
targetText = files.ConvertTextEncoding(sourceText, "iso-8859-1", "utf-8")
Response.Write "Converted text: " & targetText
Set files = Nothing
%>
```
