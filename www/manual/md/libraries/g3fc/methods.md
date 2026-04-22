# G3FC Methods

## Overview
This page summarizes every method exposed by G3FC in G3Pix AxonASP.

## Methods Reference

| Method | Returns | Description |
|---|---|---|
| Create(archivePath, sourcePaths [, password] [, options]) | Boolean | Returns True when the archive is created successfully. Returns False on invalid arguments, invalid output path resolution, empty resolved source set, or archive creation failure. |
| Extract(archivePath, outputFolder [, password]) | Boolean | Returns True when all entries are extracted successfully. Returns False on invalid arguments, invalid path resolution, or extraction failure. |
| List(archivePath [, password] [, unit] [, details]) | Array of Scripting.Dictionary or Empty | Returns a zero-based Array of Dictionary entries with Path, Size, FormattedSize, and Type keys. When details=True, each Dictionary also includes Permissions, CreationTime, and Checksum. Returns Empty when arguments are invalid, path resolution fails, or index read fails. |
| Info(archivePath, outputFilePath [, password]) | Boolean | Returns True when archive metadata is exported to the output file. Returns False on invalid arguments, invalid path resolution, or export failure. |
| Find(archivePath, pattern [, password] [, useRegex]) | Array of Scripting.Dictionary or Empty | Returns a zero-based Array of Dictionary entries for matches, each with Path and Size keys. Returns Empty when arguments are invalid, path resolution fails, or index read fails. |
| ExtractSingle(archivePath, entryPath, outputFilePath [, password]) | Boolean | Returns True when one specific entry is extracted successfully. Returns False on invalid arguments, invalid path resolution, missing entry, or extraction failure. |

## Remarks
- Method names are case-insensitive.
- ExtractSingle also accepts aliases extract-single and extract_single.
- Array-returning methods return Empty on validation/read failures and do not return partial scalar fallback values.
