# Methods

## Overview
This page lists methods exposed by the G3TAR library.

## Method List
- Create: Initializes a new TAR archive on disk.
- Open: Opens an existing TAR archive for reading or extracting.
- AddFile: Appends a single file from the disk to a newly created archive.
- AddFiles: Appends multiple files from a list or dictionary to the archive.
- AddFolder: Recursively appends an entire directory to the archive.
- AddText: Directly appends raw string content as a named file in the archive.
- ExtractAll: Unpacks the entire archive to a specific disk location.
- ExtractFile: Unpacks a single named entry from the archive to a specific local path.
- List: Generates an array containing the names of all entries inside the archive.
- GetInfo: Retrieves a Scripting.Dictionary of metadata describing a specific entry in the archive.
- Close: Flushes operations and releases any file handles locked by the archive context.

## Remarks
- Method names are case-insensitive.
- Standard read/write methods return Boolean values indicating the success or failure of the requested operation.
