# G3AXON.FUNCTIONS Method Reference

## Overview

This page lists every method exposed by the `G3AXON.FUNCTIONS` library. Methods are grouped by functional category. All method names are case-insensitive.

---

## System and Environment

| Method | Returns | Description |
|---|---|---|
| `AxCacheDirPath()` | String | Full path to the `.temp/cache/` directory with trailing separator. |
| `AxChangeDir(path)` | Boolean | Changes the process working directory. Returns `True` on success. |
| `AxChangeMode(path, mode)` | Boolean | Changes the file mode using an octal string (e.g., `"0644"`). |
| `AxChangeOwner(path, uid, gid)` | Boolean | Changes file owner and group IDs. |
| `AxChangeTimes(path, accessTime, modifyTime)` | Boolean | Sets file access and modification timestamps from Unix seconds. |
| `AxClearEnvironment()` | Boolean | Clears all process environment variables. Always returns `True`. |
| `AxCreateLink(sourcePath, linkPath)` | Boolean | Creates a hard link from source to destination. |
| `AxCurrentDir()` | String | Current process working directory path. |
| `AxCurrentUser()` | String | Current operating system user name. |
| `AxDirSeparator()` | String | OS directory separator (`\` on Windows, `/` on Unix). |
| `AxEffectiveUserId()` | Integer | Effective user ID on Unix, or `-1` on Windows. |
| `AxEngineName()` | String | AxonASP engine name string. |
| `AxEnvironmentList()` | Array | All environment entries as `KEY=VALUE` strings. |
| `AxEnvironmentValue(name [, default])` | String | Value of an environment variable, with optional default fallback. |
| `AxExecutablePath()` | String | Absolute path of the running AxonASP executable. |
| `AxExecute(command)` | String/Boolean | Runs a shell command; returns combined stdout/stderr or `False` on empty command. |
| `AxGetEnv(name)` | String | Value of an environment variable by name. |
| `AxHostnameValue()` | String | Current machine host name. |
| `AxIsPathSeparator(char)` | Boolean | `True` if the single character is a valid path separator on the current platform. |
| `AxPathListSeparator()` | String | OS path-list separator (`;` on Windows, `:` on Unix). |
| `AxProcessId()` | Integer | Current process ID (PID). |
| `AxRuntimeInfo()` | String | Multi-section plain-text diagnostic report. |
| `AxShutdownAxonASPServer()` | Boolean | Terminates the server process if shutdown is enabled in configuration. |
| `AxSystemInfo([mode])` | String | System information: OS, hostname, architecture, Go version, or all combined. |
| `AxUserConfigDirPath()` | String | Resolved path to `config/axonasp.toml`. |
| `AxUserHomeDirPath()` | String | Current user home directory path. |
| `AxVersion()` | String | AxonASP runtime version string. |

---

## Configuration

| Method | Returns | Description |
|---|---|---|
| `AxGetConfig(key)` | String | Reads a configuration key from `axonasp.toml` with optional environment override. |
| `AxGetConfigKeys()` | Array | All available configuration key names. |
| `AxGetDefaultCss()` | String | Raw content of the CSS file configured at `axfunctions.ax_default_css_path`. |
| `AxGetLogo()` | String | Configured logo as a Base64 data URI. |
| `AxPoweredByImage()` | String | Built-in "Powered by AxonASP" image as a Base64 data URI. |

---

## Math and Numeric

| Method | Returns | Description |
|---|---|---|
| `AxCeil(number)` | Double | Rounds a value up to the nearest integer boundary. |
| `AxCount(value)` | Integer | Number of elements in a VBScript array. Returns `0` for non-arrays. |
| `AxFloor(number)` | Double | Rounds a value down to the nearest integer boundary. |
| `AxFloatPrecisionDigits()` | Integer | Standard float precision digit count (`15`). |
| `AxIntegerMax()` | Integer | Maximum native 64-bit integer value (`9,223,372,036,854,775,807`). |
| `AxIntegerMin()` | Integer | Minimum native 64-bit integer value (`-9,223,372,036,854,775,808`). |
| `AxIntegerSizeBytes()` | Integer | Native integer size in bytes (`4` on 32-bit, `8` on 64-bit). |
| `AxMax(n1, n2, ...)` | Double | Largest value from the supplied arguments. |
| `AxMin(n1, n2, ...)` | Double | Smallest value from the supplied arguments. |
| `AxNumberFormat(n, decimals, decPoint, thousandsSep)` | String | Formats a number with configurable separators. |
| `AxPi()` | Double | Mathematical constant π (`3.141592653589793`). |
| `AxPlatformBits()` | Integer | Native integer size in bits (`32` or `64`). |
| `AxRand([max] / [min, max])` | Integer | Random integer within the specified range. |
| `AxSmallestFloatValue()` | Double | Smallest non-zero positive `float64` value (`5e-324`). |

---

## Array

| Method | Returns | Description |
|---|---|---|
| `AxArrayReverse(arr)` | Array | New array with elements in reverse order. |
| `AxExplode(delimiter, str [, limit])` | Array | Splits a string into an array by delimiter. |
| `AxImplode(glue, arr)` | String | Joins array elements into a string with a separator. |
| `AxRange(start, end [, step])` | Array | Numeric array from `start` to `end` with optional `step`. |

---

## String

| Method | Returns | Description |
|---|---|---|
| `AxNl2Br(str)` | String | Replaces line breaks (CRLF, LF, CR) with HTML `<br>` tags. |
| `AxPad(str, length, padStr, padType)` | String | Pads a string to a target length on the left, right, or both sides. |
| `AxRepeat(str, count)` | String | Repeats a string a specified number of times. |
| `AxStringGetCsv(str [, delimiter])` | Array | Parses one CSV row and returns the field values as an array. |
| `AxStringReplace(search, replacement, subject)` | String | Replaces all occurrences of a substring. |
| `AxTrim(str [, chars])` | String | Trims whitespace or a custom character set from both ends. |
| `AxUcFirst(str)` | String | Uppercases the first character of a string. |
| `AxW(text)` | Empty | Writes HTML-escaped text to the response output stream. |
| `AxWordCount(str [, format])` | Integer/Array | Counts words in a string, or returns them as an array. |

---

## Hash and Encoding

| Method | Returns | Description |
|---|---|---|
| `AxBase64Decode(str)` | String | Decodes a Base64 string. Returns empty string on invalid input. |
| `AxBase64Encode(str)` | String | Encodes a string to Base64. |
| `AxHash(algo, str)` | String | Hexadecimal hash using `"md5"`, `"sha1"`, or `"sha256"`. |
| `AxHexToRgb(hex)` | String | Converts `#RRGGBB` or `#RGB` to `rgb(R,G,B)`. |
| `AxHtmlSpecialChars(str)` | String | Escapes `&`, `<`, `>`, and `"` to HTML entities. |
| `AxMD5(str)` | String | 32-character lowercase MD5 hash. |
| `AxRawUrlDecode(str)` | String | URL-decodes a string, converting `+` to space first. |
| `AxRgbToHex(r, g, b)` | String | Converts RGB values to uppercase `#RRGGBB` hex. |
| `AxSHA1(str)` | String | 40-character lowercase SHA-1 hash. |
| `AxStripTags(str)` | String | Removes all HTML and XML tags from a string. |
| `AxUrlDecode(str)` | String | Percent-decodes a URL-encoded string (RFC 3986). |

---

## Validation and Type Inspection

| Method | Returns | Description |
|---|---|---|
| `AxCtypeAlnum(str)` | Boolean | `True` if all characters are ASCII alphanumeric. Empty string returns `False`. |
| `AxCtypeAlpha(str)` | Boolean | `True` if all characters are ASCII letters. Empty string returns `False`. |
| `AxEmpty(value)` | Boolean | `True` if value is `Empty`, `Null`, `""`, `0`, `0.0`, or `False`. |
| `AxFilterValidateEmail(email)` | Boolean | `True` if the string is a syntactically valid email address. |
| `AxFilterValidateIp(ip)` | Boolean | `True` if the string is a valid IPv4 or IPv6 address. |
| `AxIsFloat(value)` | Boolean | `True` if the VM internal type is `VTDouble`. |
| `AxIsInt(value)` | Boolean | `True` if the VM internal type is `VTInteger`. |
| `AxIsSet(value)` | Boolean | `True` if value is not `Empty` or `Null`. |

---

## Date, Time, and Network

| Method | Returns | Description |
|---|---|---|
| `AxDate(format [, timestamp])` | String | Formats a Unix timestamp using PHP-compatible tokens. |
| `AxGenerateGuid()` | String | RFC 4122 version 4 GUID string. |
| `AxGetRemoteFile(url)` | String/Boolean | HTTP GET response body on success, or `False` on failure. |
| `AxLastModified()` | Integer | Unix timestamp of the current script file's last modification. |
| `AxTime()` | Integer | Current Unix timestamp in seconds. |

---

## Remarks

- All method names are case-insensitive.
- Instantiate the library with `Server.CreateObject("G3AXON.FUNCTIONS")` before calling any method.
- Methods that perform file system or OS operations may behave differently on Windows versus Unix-like platforms. Refer to individual method pages for platform notes.
