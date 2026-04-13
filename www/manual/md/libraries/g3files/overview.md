# Use the G3FILES Library

## Overview
The **G3FILES** library provides high-performance file system operations for G3Pix AxonASP applications. It offers a comprehensive set of methods for reading, writing, and managing files and directories with built-in support for multiple text encodings (UTF-8, UTF-16, ISO-8859-1, ASCII) and line ending normalization. The library is optimized for server-side file management and operates within the AxonASP sandboxed file system.

## Syntax
To instantiate the library, use the following syntax:
```asp
Set files = Server.CreateObject("G3FILES")
```

## Prerequisites
- **File System Permissions**: The AxonASP service must have appropriate read and write permissions for the target directories.
- **Sandbox Configuration**: All file paths are resolved relative to the configured AxonASP sandbox root for security.

## How it Works
The G3FILES object provides direct access to the underlying host file system through the AxonASP Virtual Machine. 
- **Encoding Support**: The library automatically handles Byte Order Marks (BOM) when reading files and allows explicit BOM inclusion when writing.
- **Line Endings**: Line endings can be normalized to Windows (CRLF), Linux/Unix (LF), or legacy Mac (CR) styles during write or conversion operations.
- **Path Resolution**: Relative paths are automatically resolved within the AxonASP sandbox, ensuring that file operations cannot escape the intended application directory.

## API Reference

### Methods
- **Append**: Appends text content to an existing file or creates it if it does not exist.
- **ConvertFileEncoding**: Converts the text encoding and line endings of a file on disk.
- **ConvertTextEncoding**: Returns a string converted from one text encoding to another.
- **Copy**: Copies a file from a source path to a destination path.
- **Delete**: Permanently removes a file or directory from the file system.
- **Exists**: Returns a Boolean indicating whether a file or directory exists at the specified path.
- **List**: Returns an Array containing the names of all files within a specified directory.
- **Mkdir**: Creates a new directory or a full directory path recursively.
- **Move**: Renames or moves a file or directory to a new location.
- **NormalizeEOL**: Returns a string with line endings converted to a specified style.
- **Read**: Returns the full text content of a file using a specified encoding.
- **Size**: Returns the size of a file in bytes.
- **Write**: Writes text content to a file, overwriting any existing content.

## Code Example
The following example demonstrates how to check for a file's existence, read its content, and write a modified version to a new location.

```asp
<%
Dim files, content, success
Set files = Server.CreateObject("G3FILES")

' Check if a configuration file exists
If files.Exists("/config/settings.txt") Then
    ' Read content using UTF-8 encoding
    content = files.Read("/config/settings.txt", "utf-8")
    
    ' Append a timestamp to the content
    content = content & vbCrLf & "Last Access: " & Now()
    
    ' Write the updated content to a backup file
    success = files.Write("/backups/settings_backup.txt", content, "utf-8", "windows", True)
    
    If success Then
        Response.Write "Backup created successfully."
    Else
        Response.Write "Failed to create backup."
    End If
Else
    Response.Write "Source file not found."
End If

Set files = Nothing
%>
```
