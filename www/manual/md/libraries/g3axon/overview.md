# Use the G3AXON.FUNCTIONS Library

## Overview
The **G3AXON.FUNCTIONS** library provides a comprehensive set of native utility functions for G3Pix AxonASP applications. These functions extend the standard VBScript capabilities with high-performance system operations, environment management, advanced math routines, and string manipulation helpers. The library is optimized for zero-allocation performance and provides direct access to the underlying Go runtime and host operating system features.

## Syntax
To instantiate the library, use the following syntax:
```asp
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")
```

## Prerequisites
No external dependencies are required. The G3AXON.FUNCTIONS library is a built-in native component of the AxonASP environment.

## How it Works
The G3AXON.FUNCTIONS object acts as a gateway to specialized system and utility routines. 
- **System & Environment**: Provides access to process IDs, hostname, environment variables, and directory management.
- **Math & Numeric**: Includes advanced functions like **AxMax**, **AxMin**, and high-precision random number generation.
- **Array & String**: Offers robust tools for data manipulation, such as **AxExplode**, **AxImplode**, and **AxPad**.
- **Output**: The **Axw** method provides an HTML-escaped alternative to standard output routines.

Most functions in this library are also available as global built-ins (prefixed with "Ax") for developer convenience.

## API Reference
This library contains over 80 specialized methods. For detailed information on specific members, please refer to the individual method documentation pages.

## Code Example
The following example demonstrates how to retrieve system information and format a number using the library.

```asp
<%
Dim ax, sysInfo, formatted
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

' Retrieve a summary of system information
sysInfo = ax.AxSystemInfo("a")
Response.Write "System: " & sysInfo & "<br>"

' Format a numeric value with thousands separators
formatted = ax.AxNumberFormat(1234567.89, 2, ".", ",")
Response.Write "Formatted: " & formatted

Set ax = Nothing
%>
```
