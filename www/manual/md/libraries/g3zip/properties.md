# G3ZIP Properties

## Overview
This page provides a summary of the properties available in the **G3ZIP** library for inspecting archive state.

## Property List

- **Count**: Read-only. Returns an **Integer** representing the total number of files in the opened archive.
- **Mode**: Read-only. Returns a **String** ("r" for Read, "w" for Write) indicating the current operating mode.
- **Path**: Read-only. Returns a **String** containing the absolute physical path of the active archive.

## Remarks
- Properties are read-only and cannot be modified directly.
- The **Count** property is only valid when an archive is opened in Read mode.
- Use the **methods.md** page for actions related to these properties.
