# G3IMAGE Properties

## Overview

This page lists the properties exposed by `G3IMAGE` and the `Persits.Jpeg` compatibility layer.

## Properties

| Property | Access | Type | Description |
|---|---|---|---|
| `HasContext` | Read-only | Boolean | Indicates whether an active drawing context exists. |
| `Width` | Read/Write | Integer | Pixel width of the active context. Setting it resizes the image. |
| `Height` | Read/Write | Integer | Pixel height of the active context. Setting it resizes the image. |
| `LastError` | Read-only | String | Latest operation error message. |
| `LastMimeType` | Read-only | String | MIME type of the most recent rendered output. |
| `LastTempFile` | Read-only | String | Temporary file path used by the most recent temp render operation. |
| `LastBytes` | Read-only | Array | Byte array from the most recent render operation. |
| `DefaultFormat` | Read/Write | String | Default render format (`png`, `jpg`, or `jpeg`). |
| `JPGQuality` | Read/Write | Integer | JPEG quality setting clamped to the range `1..100`. |
| `AlignLeft` | Read-only | Integer | Drawing constant for left text alignment. |
| `AlignCenter` | Read-only | Integer | Drawing constant for centered text alignment. |
| `AlignRight` | Read-only | Integer | Drawing constant for right text alignment. |
| `FillRuleWinding` | Read-only | Integer | Drawing constant for winding fill rule. |
| `FillRuleEvenOdd` | Read-only | Integer | Drawing constant for even-odd fill rule. |
| `LineCapRound` | Read-only | Integer | Drawing constant for round line caps. |
| `LineCapButt` | Read-only | Integer | Drawing constant for butt line caps. |
| `LineCapSquare` | Read-only | Integer | Drawing constant for square line caps. |
| `LineJoinRound` | Read-only | Integer | Drawing constant for round line joins. |
| `LineJoinBevel` | Read-only | Integer | Drawing constant for bevel line joins. |
| `OriginalWidth` | Read-only | Integer | **(Persits.Jpeg Compatibility)** Original width of the opened image. |
| `OriginalHeight` | Read-only | Integer | **(Persits.Jpeg Compatibility)** Original height of the opened image. |
| `Quality` | Read/Write | Integer | **(Persits.Jpeg Compatibility)** JPEG compression quality (alias to JPGQuality). |
| `Interpolation` | Read/Write | Integer | **(Persits.Jpeg Compatibility)** Resampling algorithm (1=Nearest, 2=Bilinear, 3=Bicubic). |
| `Canvas` | Read-only | Object | **(Persits.Jpeg Compatibility)** Nested sub-object containing methods to draw text and shapes. |
| `Canvas.Pen.Color` | Read/Write | String/Integer | **(Persits.Jpeg Compatibility)** Drawing stroke color. |
| `Canvas.Pen.Width` | Read/Write | Double | **(Persits.Jpeg Compatibility)** Drawing stroke thickness. |
| `Canvas.Font.Color` | Read/Write | String/Integer | **(Persits.Jpeg Compatibility)** Font drawing color. |
| `Canvas.Font.Family` | Read/Write | String | **(Persits.Jpeg Compatibility)** Font family name or TTF path. |
| `Canvas.Font.Size` | Read/Write | Double | **(Persits.Jpeg Compatibility)** Font size in points. |
| `Canvas.Font.Bold` | Read/Write | Boolean | **(Persits.Jpeg Compatibility)** Font bold styling. |
| `Canvas.Font.Italic` | Read/Write | Boolean | **(Persits.Jpeg Compatibility)** Font italic styling. |

## Remarks

- Instantiate the library with `Server.CreateObject("G3IMAGE")` or `Server.CreateObject("Persits.Jpeg")`.
- Property names are case-insensitive.
