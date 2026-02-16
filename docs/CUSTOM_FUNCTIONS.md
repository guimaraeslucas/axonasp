# G3 AxonASP - Custom Functions Documentation

## Overview
Custom functions have been implemented in a dedicated file following VBScript conventions. All runtime-facing functions are named with the **Ax** prefix and use **PascalCase**.

## File Structure
- **Location**: `server/custom_functions.go`
- **Integration**: Functions are automatically registered in `executor.go` via `evalCustomFunction()`
- **Fallback**: Custom functions are checked first, then built-in functions

## Function Categories

### 1. Document Functions
- **Document.Write**: HTML-safe encoded version of Response.Write
  ```vb
  Document.Write "<script>alert('xss')</script>"
  ' Outputs: &lt;script&gt;alert(&#39;xss&#39;)&lt;/script&gt;
  ```

### 2. System Functions

#### AxGetEnv
Return the value of an OS environment variable
```vb
Dim pathValue
pathValue = AxGetEnv("PATH")
Response.Write pathValue
```

### 3. Array Functions

#### AxArrayMerge
Merge multiple arrays into a single contiguous array
```vb
Dim arr1, arr2, merged
arr1 = Array(1, 2, 3)
arr2 = Array(4, 5, 6)
merged = AxArrayMerge(arr1, arr2)  ' [1,2,3,4,5,6]
```

#### AxArrayContains
Search for exact value in collection/array (like in_array)
```vb
Dim fruits
fruits = Array("apple", "banana", "orange")
If AxArrayContains("banana", fruits) Then
    Response.Write "Found!"
End If
```

#### AxArrayMap
Transform array elements using callback function
```vb
' Requires callback function defined in ASP
results = AxArrayMap("FunctionName", inputArray)
```

#### AxArrayFilter
Filter array using callback function
```vb
filtered = AxArrayFilter("FilterFunction", inputArray)
```

#### AxCount
Return array/collection length (returns 0 for empty/null)
```vb
Dim arr
arr = Array("a", "b", "c")
Response.Write AxCount(arr)  ' 3
Response.Write AxCount(Empty)  ' 0
```

#### AxExplode
Split string by delimiter (with limit parameter support)
```vb
Dim parts
parts = AxExplode(",", "one,two,three,four")
Response.Write AxImplode("|", parts)  ' one|two|three|four
```

```vb
Dim nums, reversed
```vb
Dim numbers
numbers = AxRange(1, 10)  ' [1,2,3,...,10]
numbers = AxRange(1, 10, 2)  ' [1,3,5,7,9]
```

#### AxImplode
Join array elements with glue string
```vb
Dim words
words = Array("Hello", "World", "ASP")
Response.Write AxImplode(" ", words)  ' Hello World ASP
```

### 4. String Functions

#### AxStringReplace
Replace search string with replacement (supports arrays)
```vb
Dim text
text = "The fox jumps"
text = AxStringReplace("fox", "cat", text)  ' The cat jumps
```

#### AxSprintf
C-style string formatting (%s, %d, %f, %x, %X)
```vb
Dim formatted
formatted = AxSprintf("User: %s, Age: %d", "John", 25)
' Result: User: John, Age: 25
```

#### AxPad
Pad string to length with pad string
```vb
' AxPad(string, length, pad_string, pad_type)
' pad_type: 0=left, 1=right (default), 2=both
Dim padded
padded = AxPad("5", 5, "0", 0)  ' "00005"
```

#### AxRepeat
Repeat string multiple times
```vb
Response.Write AxRepeat("*", 10)  ' **********
```

#### AxUcFirst
Uppercase first character
```vb
Response.Write AxUcFirst("hello world")  ' Hello world
```

#### AxWordCount
Count words in string (or return array of words)
```vb
Dim count, words
count = AxWordCount("The quick brown fox", 0)  ' 4
words = AxWordCount("The quick brown fox", 1)  ' Array of words
```

#### AxNewLineToBr
Convert newlines to `<br>` tags
```vb
Dim text, html
text = "Line 1" & vbCrLf & "Line 2"
html = AxNewLineToBr(text)
```

#### AxTrim
Trim characters from string (default: whitespace)
```vb
Response.Write "'" & AxTrim("  hello world  ") & "'"  ' 'hello world'
```

#### AxStringGetCsv
Parse CSV string
```vb
Dim values
values = AxStringGetCsv("col1,col2,col3")
```

### 5. Math Functions

#### AxCeil, AxFloor
Round up/down
```vb
Response.Write AxCeil(4.3)  ' 5
Response.Write AxFloor(4.8)  ' 4
```

#### AxMax, AxMin
Return maximum/minimum value (accepts multiple arguments)
```vb
Response.Write AxMax(5, 12, 3, 8)  ' 12
Response.Write AxMin(5, 12, 3, 8)  ' 3
```

#### AxRand
Random integer
```vb
Response.Write AxRand(1, 10)  ' Random 1-10
```

#### AxNumberFormat
Format number with thousands separator and decimals
```vb
' AxNumberFormat(number, decimals, dec_point, thousands_sep)
Response.Write AxNumberFormat(1234567.89, 2, ".", ",")
' 1,234,567.89
```

### 6. Type Checking Functions

#### AxIsInt, AxIsFloat
Check if value is integer or float
```vb
Response.Write AxIsInt(5)  ' True
Response.Write AxIsFloat(5.5)  ' True
```

#### AxCTypeAlpha, AxCTypeAlnum
Check if all characters are alphabetic or alphanumeric
```vb
Response.Write AxCTypeAlpha("hello")  ' True
Response.Write AxCTypeAlnum("hello123")  ' True
```

#### AxEmpty
Check if value is empty (includes: Null, Empty, "", 0, False, empty arrays)
```vb
Response.Write AxEmpty("")  ' True
Response.Write AxEmpty(0)  ' True
```

#### AxIsset
Check if variable is set (not null/empty)
```vb
Response.Write AxIsset(someVariable)
```

### 7. Date/Time Functions

#### AxTime
Return current Unix timestamp
```vb
Response.Write AxTime  ' 1705432800 (example)
```

#### AxDate
Format date/time (like PHP date())
```vb
' AxDate(format, [timestamp])
Response.Write AxDate("Y-m-d")  ' 2024-01-16
Response.Write AxDate("Y-m-d H:i:s")  ' 2024-01-16 14:30:45
```

### 8. Hashing & Encoding Functions

#### AxMd5, AxSha1, AxHash
Hash string
```vb
Dim password
password = "secret"
Response.Write AxMd5(password)  ' 5EB63BBBE01EEED093CB22BB8F5ACDC3
Response.Write AxSha1(password)  ' E5FA44F2B31C1FB553B6021E7ABED18E
Response.Write AxHash("sha256", password)
```

#### AxBase64Encode, AxBase64Decode
Base64 encoding/decoding
```vb
Dim encoded, decoded
encoded = AxBase64Encode("Hello, World!")
decoded = AxBase64Decode(encoded)
```

#### AxUrlDecode, AxRawUrlDecode
Decode URL-encoded strings
```vb
Response.Write AxUrlDecode("Hello%20World%21")  ' Hello World!
```

#### AxRgbToHex
Convert RGB to hex color
```vb
Response.Write AxRgbToHex(255, 128, 0)  ' #FF8000
```

#### AxHtmlSpecialChars
Encode HTML special characters
```vb
Response.Write AxHtmlSpecialChars("<p>Test & Demo</p>")
' &lt;p&gt;Test &amp; Demo&lt;/p&gt;
```

#### AxStripTags
Remove HTML tags
```vb
Response.Write AxStripTags("<p>Hello <b>World</b></p>")
' Hello World
```

### 9. Validation Functions

#### AxFilterValidateIp
Validate IP address
```vb
If AxFilterValidateIp("192.168.1.1") Then
    Response.Write "Valid IP"
End If
```

#### AxFilterValidateEmail
Validate email address
```vb
If AxFilterValidateEmail("user@example.com") Then
    Response.Write "Valid Email"
End If
```

### 10. Request Array Functions

#### AxGetRequest, AxGetGet, AxGetPost
Get request parameters as Dictionary-like objects
```vb
Dim params
params = AxGetRequest()  ' Merged GET and POST
params = AxGetGet()  ' Only GET
params = AxGetPost()  ' Only POST
```

### 11. Include Functions

#### AxInclude
Execute an ASP file and output its result (same behavior as `<!--# include -->` directive)

```vb
AxInclude "header.asp"        ' Relative to current directory
AxInclude "./header.asp"      ' Explicit relative path
AxInclude "../shared/nav.asp" ' Parent directory
AxInclude "/includes/menu.asp" ' Virtual path from web root
```

**Path Resolution Rules:**
- **Absolute paths** (starting with `/`, `./`, `../`, or drive letter): Resolved according to their prefix
  - `/path` - Virtual path from web root
  - `./path` - Relative to current directory
  - `../path` - Relative to parent of current directory
  - `C:\path` - Windows absolute path
- **Relative paths** (no prefix): Resolved relative to current file's directory

**Security:**
- ⚠️ **Security Warning**: If you include files outside the web root directory, user may gain access to system files
- ⚠️ Local file paths only - remote URLs are NOT supported (use `AxGetRemoteFile` instead)
- Paths are validated to prevent access outside web root; errors are printed to console

**Return Value:**
- `true` - File executed successfully
- `false` - File not found or execution error (error message printed to console)

```vb
' Example usage in ASP file
<%
If AxInclude("/includes/header.asp") Then
    Response.Write "Header included successfully"
Else
    Response.Write "Failed to include header"
End If
%>

' Always execute a file and output result
<%
AxInclude "content.asp"
%>
```

**Return Value From Execution:**
When AxInclude executes the file, any output from that file is written to the Response.

#### AxIncludeOnce
Execute an ASP file only once per page execution (prevents duplicate inclusion)

```vb
AxIncludeOnce "config.asp"    ' First call executes it
AxIncludeOnce "config.asp"    ' Subsequent calls are ignored
```

**Path Resolution:**
- Same rules as AxInclude

**Security:**
- Same security restrictions as AxInclude
- Paths outside web root are rejected

**Return Value:**
- `true` - File executed successfully or already included (not an error)
- `false` - File not found or execution error (error message printed to console)

**Behavior:**
- Identical to AxInclude but tracks included files by normalized path
- Each unique file path can only be executed once per page execution
- Subsequent calls with same path are ignored but return `true` (not an error)
- Useful for including configuration or initialization files that should load only once

```vb
' Example: Include dependencies only once
<%
Function IncludeIfNeeded()
    AxIncludeOnce "/config/settings.asp"
    AxIncludeOnce "/lib/helpers.asp"
End Function

IncludeIfNeeded() ' Loads both files
IncludeIfNeeded() ' Files already loaded, calls are silently ignored
%>
```

#### AxGetRemoteFile
Fetch content from a remote URL (plain text, NOT executed as ASP)

```vb
Dim content
content = AxGetRemoteFile("https://example.com/data.txt")

If IsArray(content) Or content = False Then
    Response.Write "Failed to fetch remote file"
Else
    Response.Write "Retrieved: " & content
End If
```

**Supported Protocols:**
- `http://` - HTTP protocol
- `https://` - HTTPS protocol (secure)

**Security:**
- ⚠️ Remote URLs only - local file paths are NOT supported for security
- ⚠️ Content is returned as plain text, NOT executed as ASP code
- Cannot access `localhost` or `127.0.0.1` for security
- 10-second timeout to prevent hanging on slow connections

**Return Value:**
- `string` - File content successfully retrieved
- `false` - Failed to fetch (error message printed to console)

**Error Cases:**
- Invalid protocol (only `http://` and `https://` supported)
- Local file paths (security restriction)
- HTTP error responses (non-200 status codes)
- Network timeout or connection failure
- Invalid URL format

```vb
' Example: Fetch remote JSON (as text)
<%
Dim jsonData
jsonData = AxGetRemoteFile("https://api.example.com/data.json")

If jsonData <> False Then
    ' jsonData contains the raw JSON text
    Response.ContentType = "application/json"
    Response.Write jsonData
Else
    Response.Write "Failed to fetch API data"
End If
%>
```

**Important Notes:**
- Content is returned as plain text string
- No execution or parsing of ASP code from remote source
- Use G3JSON to parse JSON content if needed
- Consider rate limiting and timeout handling for production use

### 12. Utility Functions

#### AxGenerateGuid
Generate a new GUID
```vb
Response.Write AxGenerateGuid()
' Example: a1b2c3d4-e5f6-4a7b-8c9d-e0f1a2b3c4d5
```

#### AxBuildQueryString
Build URL query string from Dictionary
```vb
Dim params, queryString
Set params = CreateObject("Scripting.Dictionary")
params("name") = "John"
params("age") = 25
queryString = AxBuildQueryString(params)
' name=John&age=25
```

#### AxVarDump
Debug output variable (accepts any type)
```vb
Dim testArray
testArray = Array("hello", 123, 45.67)
Response.Write "<pre>"
AxVarDump testArray
Response.Write "</pre>"
```

## Naming Convention Summary

| Reference Function | G3 AxonASP Function |
|---|---|
| array_merge | AxArrayMerge |
| in_array | AxArrayContains |
| array_map | AxArrayMap |
| array_filter | AxArrayFilter |
| count | AxCount |
| explode | AxExplode |
| str_replace | AxStringReplace |
| sprintf | AxSprintf |
| var_dump | AxVarDump |
| str_pad | AxPad |
| str_repeat | AxRepeat |
| ucfirst | AxUcFirst |
| str_word_count | AxWordCount |
| nl2br | AxNewLineToBr |
| trim | AxTrim |
| ceil | AxCeil |
| floor | AxFloor |
| max | AxMax |
| min | AxMin |
| rand | AxRand |
| number_format | AxNumberFormat |
| is_int | AxIsInt |
| is_float | AxIsFloat |
| ctype_alpha | AxCTypeAlpha |
| ctype_alnum | AxCTypeAlnum |
| empty | AxEmpty |
| isset | AxIsset |
| time | AxTime |
| date | AxDate |
| md5 | AxMd5 |
| sha1 | AxSha1 |
| hash | AxHash |
| base64_encode | AxBase64Encode |
| base64_decode | AxBase64Decode |
| urldecode | AxUrlDecode |
| rawurldecode | AxRawUrlDecode |
| str_getcsv | AxStringGetCsv |
| filter_var (IP) | AxFilterValidateIp |
| filter_var (Email) | AxFilterValidateEmail |
| htmlspecialchars | AxHtmlSpecialChars |
| strip_tags | AxStripTags |
| rgb_to_hex | AxRgbToHex |
| $_REQUEST | AxGetRequest |
| $_GET | AxGetGet |
| $_POST | AxGetPost |
| uniqid/guid | AxGenerateGuid |
| http_build_query | AxBuildQueryString |
| include | AxInclude |
| include_once | AxIncludeOnce |
| file_get_contents (remote) | AxGetRemoteFile |

### 12. OS Functions (Go `os` equivalents)

#### AxChangeDir
Change current working directory.
```vb
ok = AxChangeDir("C:\\temp")
```

#### AxChangeMode
Change file mode/permissions.
```vb
ok = AxChangeMode("C:\\temp\\file.txt", "0644")
```

#### AxChangeOwner
Change file owner/group (may be unavailable on Windows or without privileges).
```vb
ok = AxChangeOwner("/tmp/file.txt", 1000, 1000)
```

#### AxHostNameValue
Return machine hostname.
```vb
host = AxHostNameValue()
```

#### AxChangeTimes
Change file access and modification times (Unix timestamps).
```vb
ok = AxChangeTimes("/tmp/file.txt", 1700000000, 1700000100)
```

#### AxClearEnvironment
Clear all environment variables for the current process.
```vb
Call AxClearEnvironment()
```

#### AxEnvironmentList
Return environment entries as array (`KEY=VALUE`).
```vb
env = AxEnvironmentList()
```

#### AxGetEnv / AxEnvironmentValue
Read environment values.
```vb
pathValue = AxGetEnv("PATH")
value = AxEnvironmentValue("APP_MODE", "production")
```

#### AxEffectiveUserId / AxProcessId / AxCurrentDir
Return effective user ID (or `-1` on unsupported platforms), process ID, and current directory.
```vb
uid = AxEffectiveUserId()
pid = AxProcessId()
cwd = AxCurrentDir()
```

#### AxIsPathSeparator
Check if first character is a path separator.
```vb
isSep = AxIsPathSeparator("/")
```

#### AxCreateLink
Create hard link (may require privileges on some platforms).
```vb
ok = AxCreateLink("source.txt", "target.txt")
```

#### AxUserCacheDirPath / AxUserConfigDirPath / AxUserHomeDirPath
Return standard user directories.
```vb
cacheDir = AxUserCacheDirPath()
configDir = AxUserConfigDirPath()
homeDir = AxUserHomeDirPath()
```

### 13. Runtime Functions

#### AxLastModified
Return last modification time (Unix timestamp) of current script.
```vb
ts = AxLastModified()
```

#### AxSystemInfo
Return system information, with support for modes `a`, `s`, `n`, `r`, `v`, `m`.
```vb
full = AxSystemInfo("a")
arch = AxSystemInfo("m")
```

#### AxCurrentUser
Return current process user name.
```vb
username = AxCurrentUser()
```

#### AxVersion
Return AxonASP runtime version. Supports `version_id` option as integer.
```vb
ver = AxVersion()
verId = AxVersion("version_id")
```

#### AxRuntimeInfo
Write runtime and environment diagnostic information to response output.
```vb
Call AxRuntimeInfo()
```

### 14. Runtime/Platform Information Helpers

#### AxDirSeparator / AxPathListSeparator
Return directory separator and path list separator for current OS.
```vb
dirSep = AxDirSeparator()
pathSep = AxPathListSeparator()
```

#### AxIntegerSizeBytes / AxIntegerMax / AxIntegerMin
Return integer size in bytes, maximum int, and minimum int.
```vb
sizeBytes = AxIntegerSizeBytes()
maxInt = AxIntegerMax()
minInt = AxIntegerMin()
```

#### AxFloatPrecisionDigits / AxSmallestFloatValue
Return decimal digits with safe round-trip precision for float and smallest non-zero float.
```vb
digits = AxFloatPrecisionDigits()
small = AxSmallestFloatValue()
```

#### AxPlatformBits / AxExecutablePath
Return architecture bits (`32` or `64`) and current executable path.
```vb
bits = AxPlatformBits()
binPath = AxExecutablePath()
```

## Testing

A comprehensive test file is available at:
- **File**: `www/tests/test_custom_functions.asp`
- **URL**: `http://localhost:4050/tests/test_custom_functions.asp`

Additional test page for system/runtime functions:
- **File**: `www/tests/test_custom_system_php_functions.asp`
- **URL**: `http://localhost:4050/tests/test_custom_system_php_functions.asp`

## Implementation Notes

1. **HTML Safety**: Document.Write automatically HTML-encodes output
2. **Type Coercion**: Functions handle mixed types gracefully
3. **Array Support**: Functions work with both VB arrays and Go slices
4. **Dictionary Support**: Functions work with Scripting.Dictionary objects
5. **Error Handling**: Functions return safe defaults on invalid input (no errors thrown)
6. **Performance**: Uses Go's standard library for cryptographic operations

## Integration with Executor

The custom functions are integrated in `executor.go` at line 1820:
```go
// Try custom functions first
if result, handled := evalCustomFunction(funcName, args, v.context); handled {
    return result, nil
}
// Then try built-in functions
if result, handled := evalBuiltInFunction(funcName, args, v.context); handled {
    return result, nil
}
```

This ensures custom functions are prioritized over built-in functions, allowing override capability.
