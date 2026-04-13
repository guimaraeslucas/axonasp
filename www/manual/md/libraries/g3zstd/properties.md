# G3ZSTD Properties

## Overview
This page provides a summary of the properties available in the **G3ZSTD** library for inspecting the state of compression operations.

## Property List

- **LastError**: Read-only. Returns a **String** containing the message of the most recent error.
- **Level**: Read-only. Returns an **Integer** representing the current default compression level.

## Remarks
- Properties are read-only and will raise an error if an assignment is attempted.
- To change the compression level, use the **SetLevel** method.
- **LastError** is cleared at the start of every successful operation.
