# G3FC Implementation

The `G3FC` library provides a way to create, read, and extract G3FC container archives within AxonASP. G3FC is a high-performance archive format with built-in compression, encryption, and forward error correction (FEC).

## Usage

Create the object using `Server.CreateObject`:

```vbscript
Set fc = Server.CreateObject("G3FC")
```

## Methods

### Create(outputPath, sourcePaths, [password], [options])
Creates a new G3FC archive.
- `outputPath`: Path for the output .g3fc file.
- `sourcePaths`: A single path (string) or an array of paths to include in the archive.
- `password`: (Optional) Password for AES encryption.
- `options`: (Optional) A `Scripting.Dictionary` with advanced settings:
    - `CompressionLevel`: ZSTD level (1-22). Default: 6.
    - `GlobalCompression`: Boolean. Default: false.
    - `FECLevel`: FEC percentage (0-50). Default: 0.
    - `SplitSize`: Split size (e.g., "100MB", "1GB").
- **Returns:** Boolean (Success)

### Extract(archivePath, destinationDir, [password])
Extracts all files from a G3FC archive.
- `archivePath`: Path to the .g3fc file.
- `destinationDir`: Directory to extract files into.
- `password`: (Optional) Password if encrypted.
- **Returns:** Boolean (Success)

### List(archivePath, [password], [unit], [details])
Lists files in the archive.
- `archivePath`: Path to the .g3fc file.
- `password`: (Optional) Password if encrypted.
- `unit`: (Optional) Size unit ("B", "KB", "MB", "GB"). Default: "KB".
- `details`: (Optional) Boolean. If true, returns more metadata.
- **Returns:** Array of Dictionaries containing file info (compatible with `For Each`).

### Find(archivePath, pattern, [password], [useRegex])
Finds files within an archive.
- `archivePath`: Path to the .g3fc file.
- `pattern`: String or Regex pattern to search for.
- `password`: (Optional) Password if encrypted.
- `useRegex`: (Optional) Boolean. If true, treats pattern as Regex.
- **Returns:** Array of Dictionaries.

### ExtractSingle(archivePath, filePathInArchive, destinationDir, [password])
Extracts a single file or directory from an archive.
- `archivePath`: Path of the .g3fc file.
- `filePathInArchive`: Path of the file inside the archive.
- `destinationDir`: Directory to extract into.
- `password`: (Optional) Password if encrypted.
- **Returns:** Boolean (Success)

### Info(archivePath, jsonOutputPath, [password])
Exports the archive's metadata index to a JSON file.
- **Returns:** Boolean (Success)

## Example: Creating an encrypted archive

```vbscript
<%
Set fc = Server.CreateObject("G3FC")
sources = Array("data/docs", "data/images/logo.png")

Set opts = Server.CreateObject("Scripting.Dictionary")
opts.Add "CompressionLevel", 9
opts.Add "GlobalCompression", True

If fc.Create("backups/secure_data.g3fc", sources, "secret123", opts) Then
    Response.Write "Archive created successfully!"
Else
    Response.Write "Failed to create archive."
End If
%>
```

## Example: Listing and Extracting

```vbscript
<%
Set fc = Server.CreateObject("G3FC")
archive = "backups/secure_data.g3fc"
pass = "secret123"

' List files
files = fc.List(archive, pass, "MB", True)

If Not IsEmpty(files) Then
    Response.Write "Files found:<br>"
    For Each f In files
        Response.Write "- " & f.Item("Path") & " (" & f.Item("FormattedSize") & ")<br>"
    Next
End If

' Extract single file
If fc.ExtractSingle(archive, "data/docs/report.pdf", "temp/extracted", pass) Then
    Response.Write "Report extracted!<br>"
End If

' Extract all
If fc.Extract(archive, "temp/full_restore", pass) Then
    Response.Write "Full restoration complete!"
End If
%>
```
