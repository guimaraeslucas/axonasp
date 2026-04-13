# G3FILES Methods

## Overview
This page provides a summary of the methods available in the **G3FILES** library for file system operations in AxonASP.

## Method List

- **Append**: Appends text content to a file or creates it. Returns a **Boolean**.
- **ConvertFileEncoding**: Converts the encoding and line endings of a file. Returns a **Boolean**.
- **ConvertTextEncoding**: Returns a string converted from one text encoding to another.
- **Copy**: Copies a file from a source to a destination. Returns a **Boolean**.
- **Delete**: Permanently removes a file or directory. Returns a **Boolean**.
- **Exists**: Returns a **Boolean** indicating whether a file or directory exists.
- **List**: Returns an **Array** containing the names of all files in a directory.
- **Mkdir**: Creates a new directory or full directory path. Returns a **Boolean**.
- **Move**: Renames or moves a file or directory. Returns a **Boolean**.
- **NormalizeEOL**: Returns a string with line endings converted to a specified style.
- **Read**: Returns the full text content of a file using a specified encoding.
- **Size**: Returns an **Integer** representing the size of a file in bytes.
- **Write**: Writes text content to a file, overwriting any existing content. Returns a **Boolean**.

## Remarks
- Method names are case-insensitive.
- Most methods support aliases like **ReadText**, **WriteText**, **Remove**, and **Rename**.
- All methods return specific data types that must be validated in production code.
- Paths are relative to the AxonASP sandbox root.
