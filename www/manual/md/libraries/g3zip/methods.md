# G3ZIP Methods

## Overview
This page provides a summary of the methods available in the **G3ZIP** library for archive manipulation in AxonASP.

## Method List

- **AddFile**: Includes a physical file into the current write-mode archive. Returns a **Boolean**.
- **AddFolder**: Recursively adds a directory and its contents into the archive. Returns a **Boolean**.
- **AddText**: Creates a new file inside the archive using a provided string. Returns a **Boolean**.
- **Close**: Finalizes the archive and releases all file handles. Returns a **Boolean**.
- **Create**: Creates a new ZIP file on the server and prepares it for writing. Returns a **Boolean**.
- **ExtractAll**: Unpacks all files from the current read-mode archive to a directory. Returns a **Boolean**.
- **ExtractFile**: Unpacks a specific file from the archive to a directory. Returns a **Boolean**.
- **GetInfo**: Retrieves detailed metadata for a file within the archive. Returns a **Scripting.Dictionary**.
- **List**: Returns a collection of all file names present in the archive. Returns a **VBArray**.
- **Open**: Opens an existing ZIP file for reading and inspection. Returns a **Boolean**.

## Remarks
- All method names are case-insensitive.
- Methods that modify the archive (AddFile, AddText, etc.) require the object to be in Write mode.
- Methods that inspect or extract (List, ExtractAll, etc.) require the object to be in Read mode.
