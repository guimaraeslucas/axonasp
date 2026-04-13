# Use the G3ZSTD Library

## Overview
The **G3ZSTD** library provides high-performance data compression and decompression services for G3Pix AxonASP applications using the Zstandard (Zstd) algorithm. It is optimized for both speed and compression ratio, offering a wide range of compression levels. The library supports in-memory processing of strings and byte arrays, batch processing of arrays, and streaming file compression to minimize memory footprint.

## Syntax
To instantiate the library, use the following syntax:
```asp
Set zstd = Server.CreateObject("G3ZSTD")
```

## Prerequisites
No external dependencies are required. The G3ZSTD library is a built-in native component of the G3Pix AxonASP environment.

## How it Works
The G3ZSTD object provides a stateless interface for most compression operations but maintains internal state for the **Level** configuration and the **LastError** property. 
- **Levels**: Compression levels range from -5 (fastest) to 22 (highest compression). The default level is 3.
- **Data Types**: The library seamlessly handles VBScript strings and byte arrays (VBArray of bytes).
- **File Streaming**: Methods like **CompressFile** and **DecompressFile** use streaming I/O to process large files without loading the entire content into memory, ensuring system stability.
- **Resource Management**: While the object is managed by the AxonASP garbage collector, the **Clear** method can be used to explicitly release internal encoder/decoder buffers.

## API Reference

### Methods
- **Clear**: Releases internal buffers and resets the error state.
- **Compress**: Compresses a string or byte array into a Zstd-encoded byte array.
- **CompressFile**: Compresses a file from disk and saves the result to a new file.
- **CompressMany**: Performs batch compression on an array of inputs.
- **Decompress**: Decompresses a Zstd-encoded byte array into its original byte array.
- **DecompressFile**: Decompresses a Zstd file from disk to its original state.
- **DecompressMany**: Performs batch decompression on an array of encoded inputs.
- **DecompressText**: Decompresses a Zstd-encoded byte array directly into a UTF-8 string.
- **SetLevel**: Configures the default compression level for the object instance.

### Properties
- **LastError**: Returns the description of the most recent error encountered.
- **Level**: Returns the currently configured default compression level.

## Code Example
The following example demonstrates how to compress a string and luego decompress it back to text.

```asp
<%
Dim zstd, original, compressed, recovered
Set zstd = Server.CreateObject("G3ZSTD")

' Original data
original = "AxonASP high-performance compression test data."

' Compress data at level 9
compressed = zstd.Compress(original, 9)

' Decompress directly to string
recovered = zstd.DecompressText(compressed)

If recovered = original Then
    Response.Write "Compression and Decompression Successful"
Else
    Response.Write "Operation Failed: " & zstd.LastError
End If

Set zstd = Nothing
%>
```
