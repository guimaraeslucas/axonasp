# G3ZSTD Methods

## Overview
This page provides a summary of the methods available in the **G3ZSTD** library for data compression and decompression in G3Pix AxonASP.

## Method List

- **Clear**: Releases internal encoder/decoder resources and resets the object state. Returns a **Boolean**.
- **Compress**: Encodes a string or byte array using Zstandard. Returns a **VBArray** of bytes.
- **CompressFile**: Compresses a file from a source path to a target path. Returns a **Boolean**.
- **CompressMany**: Compresses each element of an input array. Returns an **Array** of VBArrays.
- **Decompress**: Decodes Zstd-compressed bytes back to their original form. Returns a **VBArray** of bytes.
- **DecompressFile**: Decodes a Zstd-compressed file to a target path. Returns a **Boolean**.
- **DecompressMany**: Decodes each element of an input array of compressed payloads. Returns an **Array** of VBArrays.
- **DecompressText**: Decodes Zstd-compressed bytes directly to a UTF-8 string. Returns a **String**.
- **SetLevel**: Sets the default compression level for the instance. Returns a **Boolean**.

## Remarks
- All method names are case-insensitive.
- Methods that return objects or arrays must be checked for validity before use.
- Compression levels outside the range -5 to 22 will result in an error.
