# G3 AxonASP - Custom Functions Documentation

## Overview
Custom functions have been implemented in a dedicated file following VBScript conventions and PHP behavior patterns. All functions are named with the **Ax** prefix and use **PascalCase**.

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

### 2. Array Functions

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

#### AxArrayReverse
Reverse array order
```vb
Dim nums, reversed
nums = Array(1, 2, 3, 4, 5)
reversed = AxArrayReverse(nums)  ' [5,4,3,2,1]
```

#### AxRange
Create array of values from start to end
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

### 3. String Functions

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

### 4. Math Functions

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

### 5. Type Checking Functions

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

### 6. Date/Time Functions

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

### 7. Hashing & Encoding Functions

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

### 8. Validation Functions

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

### 9. Request Array Functions

#### AxGetRequest, AxGetGet, AxGetPost
Get request parameters as Dictionary-like objects
```vb
Dim params
params = AxGetRequest()  ' Merged GET and POST
params = AxGetGet()  ' Only GET
params = AxGetPost()  ' Only POST
```

### 10. Utility Functions

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

| PHP Function | G3 AxonASP Function |
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

## Testing

A comprehensive test file is available at:
- **File**: `www/test_custom_functions.asp`
- **URL**: `http://localhost:4050/test_custom_functions.asp`

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
