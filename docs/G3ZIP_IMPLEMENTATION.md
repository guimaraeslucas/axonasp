# G3ZIP Implementation

The `G3ZIP` library provides a way to create, read, and extract ZIP files within AxonASP. It is built on top of Go's standard `archive/zip` package but tailored for web usage.

## Usage

Create the object using `Server.CreateObject`:

```vbscript
Set zip = Server.CreateObject("G3ZIP")
```

## Methods

### Create(path)
Creates a new ZIP file at the specified virtual/relative path.
- **Returns:** Boolean (Success)

### Open(path)
Opens an existing ZIP file for reading/listing.
- **Returns:** Boolean (Success)

### AddFile(sourcePath, [nameInZip])
Adds a file from the disk to an open ZIP (must be in `Create` mode).
- `sourcePath`: Path to the file on disk.
- `nameInZip`: (Optional) The name/path the file should have inside the ZIP.
- **Returns:** Boolean (Success)

### AddFolder(sourcePath, [nameInZip])
Adds a folder and its contents recursively to an open ZIP (must be in `Create` mode).
- `sourcePath`: Path to the folder on disk.
- `nameInZip`: (Optional) The name/path the folder should have inside the ZIP.
- **Returns:** Boolean (Success)

### AddText(nameInZip, content)
Creates a file inside the ZIP with the given text content.
- `nameInZip`: The name/path the file should have inside the ZIP.
- `content`: The text content for the file.
- **Returns:** Boolean (Success)

### ExtractAll(targetPath)
Extracts all contents of the open ZIP to the specified directory.
- **Returns:** Boolean (Success)

### ExtractFile(fileName, targetPath)
Extracts a single file from the open ZIP to the specified directory.
- **Returns:** Boolean (Success)

### List()
Returns an array of strings containing the paths of all files inside the ZIP.
- **Returns:** Array of strings

### GetFileInfo(fileName)
Returns a `Scripting.Dictionary` containing metadata about a file inside the ZIP.
- `Name`: String
- `Size`: Number (Uncompressed size)
- `PackedSize`: Number (Compressed size)
- `Modified`: String (RFC3339 date)
- `IsDir`: Boolean
- **Returns:** Dictionary or `Nothing` if not found.

### Close()
Closes the ZIP file and finalizes writing (if in `Create` mode).
- **Returns:** Boolean

## Properties

- `Path`: (Read-only) The absolute path to the current ZIP file.
- `Mode`: (Read-only) "r" for read, "w" for write.
- `Count`: (Read-only) Number of files in the ZIP (Reader mode only).

## Example: Creating a ZIP

```vbscript
<%
Set zip = Server.CreateObject("G3ZIP")
If zip.Create("temp/my_archive.zip") Then
    zip.AddFile "data.txt", "docs/data.txt"
    zip.AddText "hello.txt", "Hello from AxonASP!"
    zip.AddFolder "images", "gallery"
    zip.Close()
    Response.Write "ZIP created successfully!"
Else
    Response.Write "Failed to create ZIP."
End If
%>
```

## Example: Extracting a ZIP

```vbscript
<%
Set zip = Server.CreateObject("G3ZIP")
If zip.Open("uploads/archive.zip") Then
    files = zip.List()
    For Each f In files
        Response.Write "File: " & f & "<br>"
    Next
    
    If zip.ExtractAll("temp/extracted") Then
        Response.Write "Extracted successfully!"
    End If
    zip.Close()
End If
%>
```
