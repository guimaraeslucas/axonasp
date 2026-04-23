# AxonASP and VBScript Error Code Reference

## Overview

This page lists all error codes recognized by G3Pix AxonASP, divided into two distinct categories:

- **AxonASP Error Codes** — Internal platform errors raised by the GoLang runtime, server, FastCGI process, CLI, service wrapper, caching layer, and built-in libraries. These are defined in `axonvm/axonvmerrorcodes.go` and are exclusive to the AxonASP engine. They are always expressed as plain decimal integers.
- **VBScript Error Codes** — Standard VBScript runtime and syntax errors, compatible with the original Microsoft VBScript specification. These are defined in `vbscript/vberrorcodes.go`. VBScript errors can be represented in multiple numeric formats depending on the context in which they appear.

## VBScript Error Number Formats

VBScript error numbers are exposed in two common formats:

- **Decimal** — The short numeric code used in `Err.Number` at the ASP/VBScript level. For example, `13` for a Type Mismatch.
- **Hexadecimal (HRESULT)** — The long form used by COM automation and reported by some host environments. Calculated as `0x800A0000 + decimal_code`. For example, error `13` becomes `0x800A000D`. This is the format most commonly seen in browser-side JavaScript `Error.description` properties and in COM-aware debugging tools.

When `Err.Number` is read in an ASP script, it always returns the short decimal value. The hexadecimal HRESULT form is informational and used for cross-platform or COM interoperability references.

---

## AxonASP Error Codes

AxonASP error codes are internal to the G3Pix AxonASP platform. They are never raised as VBScript `Err.Number` values. They appear in server logs, CLI output, HTTP error responses, and the `ASPError` intrinsic object when an AxonASP-level failure prevents normal script execution.

### HTTP Standard (400–504)

| Code | Description |
|------|-------------|
| 400 | Bad Request |
| 401 | Unauthorized |
| 403 | Forbidden |
| 404 | Not Found |
| 405 | Method Not Allowed |
| 413 | Payload Too Large |
| 414 | URI Too Long |
| 500 | Internal Server Error |
| 501 | Not Implemented |
| 502 | Bad Gateway |
| 503 | Service Unavailable |
| 504 | Gateway Timeout |

### Configuration and Setup (1000–1013)

| Code | Description |
|------|-------------|
| 1000 | Invalid configuration |
| 1001 | Invalid .env file or configuration |
| 1002 | Root directory not set |
| 1003 | Root directory invalid |
| 1004 | Warning: Root directory does not exist |
| 1005 | Port invalid |
| 1006 | Server location invalid |
| 1007 | Could not listen on specified port/address |
| 1008 | COM provider mode invalid |
| 1009 | Default pages invalid |
| 1010 | Debug invalid |
| 1011 | Invalid locale |
| 1012 | Invalid timezone |
| 1013 | Viper: Failed to read configuration file, using defaults |

### File System (2000–2011)

| Code | Description |
|------|-------------|
| 2000 | Missing file path |
| 2001 | File not found |
| 2002 | Could not read file |
| 2003 | Could not resolve current directory |
| 2004 | Path is a directory |
| 2005 | Bad file name |
| 2006 | Wrong file size |
| 2007 | Wrong file type |
| 2008 | File type not allowed |
| 2009 | Extension not allowed |
| 2010 | The selected file extension is not enabled in global.execute_as_asp |
| 2011 | Failed to read the requested ASP file |

### Runtime and Execution (3000–3011)

| Code | Description |
|------|-------------|
| 3000 | Runtime error |
| 3001 | Panic |
| 3002 | Internal GoLang panic |
| 3003 | Internal error |
| 3004 | Out of memory |
| 3005 | Memory limit exceeded |
| 3006 | Overflow |
| 3007 | Underflow |
| 3008 | Time expired |
| 3009 | Time execution error |
| 3010 | Expired |
| 3011 | Server forced to shutdown |

### Script and AxonVM (4000–4011)

| Code | Description |
|------|-------------|
| 4000 | Compile error |
| 4001 | Script timeout |
| 4002 | Function not implemented |
| 4003 | Not implemented |
| 4004 | Error on library |
| 4005 | Error on custom function |
| 4006 | AxonVM error |
| 4007 | Invalid procedure call or argument |
| 4008 | Invalid array bond/type |
| 4009 | Interactive desktop functions are not supported in ASP server-side execution |
| 4010 | Response buffer limit exceeded |
| 4011 | Script timeout reached and execution goroutine was detached |

### Caching (5000–5003)

| Code | Description |
|------|-------------|
| 5000 | Invalid cache version |
| 5001 | Invalid cache file |
| 5002 | Cache cleanup invalid |
| 5003 | Include cache max memory invalid |

### FastCGI, CLI, and Miscellaneous (6000–6008)

| Code | Description |
|------|-------------|
| 6000 | FastCGI pipe closed unexpectedly |
| 6001 | FastCGI protocol error |
| 6002 | Required CLI argument missing |
| 6003 | Invalid name |
| 6004 | This is a test |
| 6005 | CLI run command not enabled in configuration |
| 6006 | CLI not enabled in configuration |
| 6007 | CLI: Missing file path for -r option |
| 6008 | MSWC.PageCounter is disabled. Enable it in config/axonasp.toml |

### Service Wrapper (6100–6109)

| Code | Description |
|------|-------------|
| 6100 | Service wrapper failed to create service instance |
| 6101 | Service wrapper failed to create service logger |
| 6102 | Service wrapper failed while running service loop |
| 6103 | Service wrapper failed to execute control command |
| 6104 | Service wrapper failed to resolve configured executable path |
| 6105 | Service wrapper executable target was not found |
| 6106 | Service wrapper failed to start configured executable |
| 6107 | Service wrapper failed to stop child process |
| 6108 | Service wrapper detected unexpected child process termination |
| 6109 | Service wrapper found an invalid environment variable entry |

### G3FC Archive (7000–7006)

| Code | Description |
|------|-------------|
| 7000 | Invalid G3FC archive header or magic number |
| 7001 | G3FC decryption failed: incorrect password or corrupted data |
| 7002 | G3FC decompression failed |
| 7003 | G3FC checksum mismatch: file may be corrupted |
| 7004 | G3FC password required for this encrypted archive |
| 7005 | Requested file not found in G3FC archive |
| 7006 | Shutdown function called from ASP script |

### Request and ADODB.Stream Constraints (8000–8014)

| Code | Description |
|------|-------------|
| 8000 | Cannot use Request.Form after calling Request.BinaryRead |
| 8001 | Cannot call Request.BinaryRead after using Request.Form |
| 8010 | Operation is not allowed when the object is closed |
| 8011 | The stream Type property cannot be changed when Position is not zero |
| 8012 | The stream Charset property cannot be set when Position is not zero |
| 8013 | Arguments are of the wrong type, are out of range, or are in conflict with one another |
| 8014 | ADODB.Stream I/O error: file operation failed |

### G3DB Native Database Library (9000–9016)

| Code | Description |
|------|-------------|
| 9000 | G3DB: connection is already open; call Close first |
| 9001 | G3DB: connection is not open; call Open or OpenFromEnv first |
| 9002 | G3DB.Open: requires two arguments: driver and DSN |
| 9003 | G3DB.Query: requires an SQL string argument |
| 9004 | G3DB.Exec: requires an SQL string argument |
| 9005 | G3DB.Prepare: requires an SQL string argument |
| 9006 | G3DB: unsupported database driver |
| 9007 | G3DB: connection test (Ping) failed |
| 9008 | G3DB: query execution failed |
| 9009 | G3DB: exec statement failed |
| 9010 | G3DB: statement preparation failed |
| 9011 | G3DB: transaction operation failed |
| 9012 | G3DB: row scan failed |
| 9013 | G3DB: result set is already closed |
| 9014 | G3DB: prepared statement is already closed |
| 9015 | G3DB: transaction is already closed |
| 9016 | G3DB.OpenFromEnv: missing or incomplete configuration keys in axonasp.toml |

---

## VBScript Error Codes

The following error codes follow the standard VBScript specification. They are surfaced through the `Err` object at the ASP/VBScript level (`Err.Number`, `Err.Description`). The **Hex (HRESULT)** column shows the long-form COM error code, computed as `0x800A0000 + decimal`, which is the format used by some COM-aware tools and host environments.

### Runtime Errors

| Decimal | Hex (HRESULT) | Description |
|---------|---------------|-------------|
| 5 | 0x800A0005 | Invalid procedure call or argument |
| 6 | 0x800A0006 | Overflow |
| 7 | 0x800A0007 | Out of memory |
| 9 | 0x800A0009 | Subscript out of range |
| 10 | 0x800A000A | The array is of fixed length or temporarily locked |
| 11 | 0x800A000B | Division by zero |
| 13 | 0x800A000D | Type mismatch |
| 14 | 0x800A000E | Out of string space (overflow) |
| 17 | 0x800A0011 | Cannot perform the requested operation |
| 28 | 0x800A001C | Stack overflow |
| 35 | 0x800A0023 | Undefined SUB procedure or Function |
| 48 | 0x800A0030 | Error loading DLL |
| 51 | 0x800A0033 | Internal error |
| 52 | 0x800A0034 | Bad file name or number |
| 53 | 0x800A0035 | File not found |
| 54 | 0x800A0036 | Bad file mode |
| 55 | 0x800A0037 | File is already open |
| 57 | 0x800A0039 | Device I/O error |
| 58 | 0x800A003A | File already exists |
| 61 | 0x800A003D | Disk space is full |
| 62 | 0x800A003E | Input beyond the end of the file |
| 67 | 0x800A0043 | Too many files |
| 68 | 0x800A0044 | Device not available |
| 70 | 0x800A0046 | Permission denied |
| 71 | 0x800A0047 | Disk not ready |
| 74 | 0x800A004A | Cannot rename with different drive |
| 75 | 0x800A004B | Path/file access error |
| 76 | 0x800A004C | Path not found |
| 91 | 0x800A005B | Object variable not set |
| 92 | 0x800A005C | For loop is not initialized |
| 94 | 0x800A005E | Invalid use of Null |
| 322 | 0x800A0142 | Could not create the required temporary file |
| 424 | 0x800A01A8 | Could not find target object |
| 429 | 0x800A01AD | AxonASP cannot create object |
| 430 | 0x800A01AE | Class does not support Automation |
| 432 | 0x800A01B0 | File name or class name not found during Automation operation |
| 438 | 0x800A01B6 | Object doesn't support this property or method |
| 440 | 0x800A01B8 | Automation error |
| 445 | 0x800A01BD | Object does not support this action |
| 446 | 0x800A01BE | Object does not support the named arguments |
| 447 | 0x800A01BF | Object does not support the current locale |
| 448 | 0x800A01C0 | Named argument not found |
| 449 | 0x800A01C1 | Parameters are not optional |
| 450 | 0x800A01C2 | Wrong number of parameters or invalid property assignment |
| 451 | 0x800A01C3 | Is not a collection of objects |
| 453 | 0x800A01C5 | The specified DLL function was not found |
| 455 | 0x800A01C7 | Code resource lock error |
| 457 | 0x800A01C9 | This key already associated with an element of this collection |
| 458 | 0x800A01CA | Variable uses an Automation type not supported in VBScript |
| 462 | 0x800A01CE | The remote server does not exist or is not available |
| 481 | 0x800A01E1 | Image is invalid |
| 500 | 0x800A01F4 | Variable not defined |
| 501 | 0x800A01F5 | Illegal assignment |
| 502 | 0x800A01F6 | The object is not safe for scripting |
| 503 | 0x800A01F7 | Object not safe for initializing |
| 504 | 0x800A01F8 | Object cannot create a secure environment |
| 505 | 0x800A01F9 | Invalid or unqualified reference |
| 506 | 0x800A01FA | Class/Type is not defined |
| 507 | 0x800A01FB | Unexpected error |

### Syntax and Compiler Errors

| Decimal | Hex (HRESULT) | Description |
|---------|---------------|-------------|
| 1001 | 0x800A03E9 | Insufficient memory |
| 1002 | 0x800A03EA | Syntax error |
| 1003 | 0x800A03EB | Missing ':' |
| 1005 | 0x800A03ED | Missing '(' |
| 1006 | 0x800A03EE | Missing ')' |
| 1007 | 0x800A03EF | Missing ']' |
| 1010 | 0x800A03F2 | Missing identifier |
| 1011 | 0x800A03F3 | Missing '=' |
| 1012 | 0x800A03F4 | Missing 'If' |
| 1013 | 0x800A03F5 | Missing 'To' |
| 1014 | 0x800A03F6 | Missing 'End' |
| 1015 | 0x800A03F7 | Missing 'Function' |
| 1016 | 0x800A03F8 | Missing 'Sub' |
| 1017 | 0x800A03F9 | Missing 'Then' |
| 1018 | 0x800A03FA | Missing 'Wend' |
| 1019 | 0x800A03FB | Missing 'Loop' |
| 1020 | 0x800A03FC | Missing 'Next' |
| 1021 | 0x800A03FD | Missing 'Case' |
| 1022 | 0x800A03FE | Missing 'Select' |
| 1023 | 0x800A03FF | Missing expression |
| 1024 | 0x800A0400 | Missing statement |
| 1025 | 0x800A0401 | Missing end of statement |
| 1026 | 0x800A0402 | Requires an integer constant |
| 1027 | 0x800A0403 | Missing 'While' or 'Until' |
| 1028 | 0x800A0404 | Missing 'While', 'Until', or end of statement |
| 1029 | 0x800A0405 | Too many locals or arguments |
| 1030 | 0x800A0406 | Identifier too long |
| 1031 | 0x800A0407 | Invalid number |
| 1032 | 0x800A0408 | Invalid character |
| 1033 | 0x800A0409 | Unterminated string constant |
| 1034 | 0x800A040A | Unterminated comment |
| 1037 | 0x800A040D | Invalid use of 'Me' keyword |
| 1038 | 0x800A040E | 'Loop' statement is missing 'Do' |
| 1039 | 0x800A040F | Invalid 'Exit' statement |
| 1040 | 0x800A0410 | Invalid 'For' loop control variable |
| 1041 | 0x800A0411 | Name redefined |
| 1042 | 0x800A0412 | Must be the first statement on the line |
| 1043 | 0x800A0413 | Cannot be assigned to non-ByVal argument |
| 1044 | 0x800A0414 | Cannot use parentheses when calling Sub |
| 1045 | 0x800A0415 | Requires a literal constant |
| 1046 | 0x800A0416 | Missing 'In' |
| 1047 | 0x800A0417 | Missing 'Class' |
| 1048 | 0x800A0418 | Must be inside a class definition |
| 1049 | 0x800A0419 | Missing Let, Set, or Get in property declaration |
| 1050 | 0x800A041A | Missing 'Property' |
| 1051 | 0x800A041B | The number of parameters must be consistent with the attribute description |
| 1052 | 0x800A041C | Cannot have more than one default attribute/method in a class |
| 1053 | 0x800A041D | Class Initialize or Terminate does not have arguments |
| 1054 | 0x800A041E | Property Set or Let must have at least one parameter |
| 1055 | 0x800A041F | Error at 'Next' |
| 1056 | 0x800A0420 | 'Default' can only be on 'Property', 'Function', or 'Sub' |
| 1057 | 0x800A0421 | 'Default' must also specify 'Public' |
| 1058 | 0x800A0422 | Can only specify 'Default' on Property Get |

### Regular Expression Errors

| Decimal | Hex (HRESULT) | Description |
|---------|---------------|-------------|
| 5016 | 0x800A1398 | Requires a regular expression object |
| 5017 | 0x800A1399 | Regular expression syntax error |
| 5018 | 0x800A139A | The number of words error |
| 5019 | 0x800A139B | Regular expression is missing ']' |
| 5020 | 0x800A139C | Regular expression is missing ')' |
| 5021 | 0x800A139D | Character set cross-border |

### Special and Locale Errors

| Decimal | Hex (HRESULT) | Description |
|---------|---------------|-------------|
| 32766 | 0x800A7FFE | True |
| 32767 | 0x800A7FFF | False |
| 32811 | 0x800A802B | Element was not found |
| 32812 | 0x800A802C | The specified date is not available in the current locale's calendar |

## Remarks

- **AxonASP error codes** are defined in `axonvm/axonvmerrorcodes.go`. They are not accessible via `Err.Number` inside an ASP script. They surface in server log output, error page responses, and the `ASPError` object when the engine itself encounters a platform-level failure before or after script execution.
- **VBScript error codes** are defined in `vbscript/vberrorcodes.go`. They are raised during script execution and are directly accessible via `Err.Number` and `Err.Description` inside an `On Error Resume Next` block.
- The HRESULT hexadecimal form of VBScript errors follows the standard COM convention: `HRESULT = 0x800A0000 + decimal_code`. This value is what COM-aware environments and some external debugging tools report. AxonASP itself always exposes the short decimal form through `Err.Number`.
- The error codes `32766` (True) and `32767` (False) are compatibility constants and are not raised as operational errors.
