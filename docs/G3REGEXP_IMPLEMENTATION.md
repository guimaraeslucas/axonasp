## G3REGEXP Library Implementation Summary

### Overview
A comprehensive regular expression library has been implemented for AxonASP, providing professional-grade pattern matching and text manipulation capabilities using Go's powerful regexp engine for VBScript-compatible RegExp operations.

### Files Created/Modified

#### New/Modified Files
1. **`server/regexp_lib.go`** (561 lines)
   - Complete implementation of G3REGEXP library
   - Pattern matching with RegExp object
   - Global and case-insensitive modes
   - Match execution and results
   - Text replacement with patterns
   - Multiline support

#### Integration
1. **`server/executor_libraries.go`**
   - Added RegExpLibrary wrapper for ASPLibrary interface compatibility
   - Enables: `Set regex = Server.CreateObject("G3REGEXP")`
   - Also supports: `Server.CreateObject("REGEXP")`

### Key Features Implemented

✓ **RegExp Properties**
  - `Pattern` - The regular expression pattern
  - `IgnoreCase` - Case-insensitive matching (boolean)
  - `Global` - Match all occurrences (boolean)
  - `Multiline` - Multiline mode for ^ and $ anchors
  - `Source` - Alias for Pattern property

✓ **Pattern Matching**
  - `Test(text)` - Test if pattern matches (boolean)
  - `Execute(text)` - Execute pattern and return matches
  - `Replace(text, replacement)` - Replace matched text
  - Pattern compilation with flags applied
  - Error handling for invalid patterns

✓ **Match Results**
  - RegExpMatches collection
  - Match count property
  - Individual match access
  - Match value, index, and length information
  - 0-based indexing for matches

✓ **Advanced Features**
  - Capture groups support
  - Backreferences in replacement
  - Global search across entire string
  - Case-insensitive matching
  - Multiline regex anchors

### Architecture

**Class Hierarchy**:
```
Component (interface)
  └─ G3REGEXP
      ├─ Properties: Pattern, IgnoreCase, Global, Multiline
      ├─ Methods: Test(), Execute(), Replace()
      ├─ compilePattern()
      ├─ lastMatches array
      └─ Error handling

RegExpMatch
  ├─ Value (matched text)
  ├─ Index (0-based position)
  └─ Length (match length)

RegExpMatches
  ├─ Collection of RegExpMatch
  ├─ Count property
  └─ Item access
```

**Pattern Compilation**:
1. Pattern set via property or constructor
2. Flags applied (IgnoreCase, Multiline)
3. Pattern compiled to Go regexp
4. Compiled pattern cached
5. Recompiled on property change

### Usage Examples

#### Basic Pattern Test
```vbscript
Dim regex, isMatch
Set regex = Server.CreateObject("G3REGEXP")

regex.Pattern = "^\d{3}-\d{3}-\d{4}$"  ' Phone number pattern
isMatch = regex.Test("555-123-4567")

If isMatch Then
    Response.Write "Valid phone number"
End If
```

#### Case-Insensitive Matching
```vbscript
Dim regex, text
Set regex = Server.CreateObject("G3REGEXP")

regex.Pattern = "hello"
regex.IgnoreCase = True

If regex.Test("HELLO world") Then
    Response.Write "Found match (case-insensitive)"
End If
```

#### Global Pattern Matching
```vbscript
Dim regex, matches, i
Set regex = Server.CreateObject("G3REGEXP")

regex.Pattern = "\d+"  ' Find all numbers
regex.Global = True

text = "I have 2 apples, 5 oranges, and 10 grapes"
Set matches = regex.Execute(text)

If matches.Count > 0 Then
    Response.Write "Found " & matches.Count & " numbers:<br>"
    For i = 0 To matches.Count - 1
        Response.Write matches.Item(i).Value & "<br>"
    Next
End If
```

#### Email Validation
```vbscript
Dim regex, email
Set regex = Server.CreateObject("G3REGEXP")

' Email pattern
regex.Pattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"

email = Request.Form("email")

If regex.Test(email) Then
    Response.Write "Valid email address"
Else
    Response.Write "Invalid email address"
End If
```

#### Text Replacement
```vbscript
Dim regex, text, result
Set regex = Server.CreateObject("G3REGEXP")

' Replace all occurrences of word
regex.Pattern = "\btest\b"
regex.IgnoreCase = True
regex.Global = True

text = "This is a test. Testing is important for test suites."
result = regex.Replace(text, "exam")

' Result: This is a exam. Examning is important for exam suites.
Response.Write result
```

#### URL Validation
```vbscript
Dim regex, url
Set regex = Server.CreateObject("G3REGEXP")

' Simple URL pattern
regex.Pattern = "^https?:\/\/[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
regex.IgnoreCase = True

url = "https://www.example.com/page"

If regex.Test(url) Then
    Response.Write "Valid URL"
End If
```

#### Extract Numbers from String
```vbscript
Dim regex, text, matches, i
Set regex = Server.CreateObject("G3REGEXP")

regex.Pattern = "\d+"
regex.Global = True

text = "Order #12345: Cost $99.99 for item code 567890"
Set matches = regex.Execute(text)

Response.Write "Found " & matches.Count & " numbers:<br>"
For i = 0 To matches.Count - 1
    Response.Write "- " & matches.Item(i).Value & "<br>"
Next
```

#### Phone Number Formatting
```vbscript
Dim regex, phone, formatted
Set regex = Server.CreateObject("G3REGEXP")

phone = Request.Form("phone")

' Extract only digits
regex.Pattern = "\D"
regex.Global = True
phone = regex.Replace(phone, "")

' Format as (XXX) XXX-XXXX
regex.Pattern = "^(\d{3})(\d{3})(\d{4})$"
formatted = regex.Replace(phone, "($1) $2-$3")

Response.Write "Formatted: " & formatted
```

#### HTML Tag Removal
```vbscript
Dim regex, html, plain
Set regex = Server.CreateObject("G3REGEXP")

html = "<p>Hello <b>World</b>!</p>"

' Remove all HTML tags
regex.Pattern = "<[^>]*>"
regex.Global = True

plain = regex.Replace(html, "")
Response.Write plain  ' Output: Hello World!
```

#### Word Boundary Matching
```vbscript
Dim regex, text, result
Set regex = Server.CreateObject("G3REGEXP")

regex.Pattern = "\b\w{3}\b"  ' 3-letter words
regex.Global = True
regex.IgnoreCase = True

text = "The cat and dog are very cute animals"
Set matches = regex.Execute(text)

Response.Write "3-letter words found: " & matches.Count & "<br>"
For i = 0 To matches.Count - 1
    Response.Write "- " & matches.Item(i).Value & "<br>"
Next
```

#### Multiline Pattern Matching
```vbscript
Dim regex, text, matches
Set regex = Server.CreateObject("G3REGEXP")

regex.Pattern = "^[A-Z]"  ' Start of line with capital letter
regex.Global = True
regex.Multiline = True

text = "First line" & vbCrLf & "Second line" & vbCrLf & "Third line"
Set matches = regex.Execute(text)

Response.Write "Lines starting with capital: " & matches.Count
```

#### CSV Field Extraction
```vbscript
Dim regex, csvLine, fields, i
Set regex = Server.CreateObject("G3REGEXP")

csvLine = """John Smith"",30,""New York"",""john@example.com"""

' Match quoted or unquoted fields
regex.Pattern = """([^""]*)""|([^,]*)"
regex.Global = True

Set fields = regex.Execute(csvLine)

For i = 0 To fields.Count - 1
    Response.Write fields.Item(i).Value & "<br>"
Next
```

#### Whitespace Normalization
```vbscript
Dim regex, text, normalized
Set regex = Server.CreateObject("G3REGEXP")

text = "This   has  multiple    spaces"

' Replace multiple spaces with single space
regex.Pattern = " +"
regex.Global = True

normalized = regex.Replace(text, " ")
Response.Write normalized  ' Output: This has multiple spaces
```

### RegExp Pattern Reference

#### Character Classes
```
.           Match any character except newline
\d          Match digit [0-9]
\D          Match non-digit
\w          Match word character [a-zA-Z0-9_]
\W          Match non-word character
\s          Match whitespace
\S          Match non-whitespace
[abc]       Match a, b, or c
[a-z]       Match lowercase letters
[^abc]      Match anything except a, b, c
```

#### Anchors
```
^           Start of string (or line in multiline)
$           End of string (or line in multiline)
\b          Word boundary
\B          Non-word boundary
```

#### Quantifiers
```
*           0 or more
+           1 or more
?           0 or 1
{n}         Exactly n
{n,}        n or more
{n,m}       Between n and m
```

#### Groups and Backreferences
```
(abc)       Capture group
(?:abc)     Non-capturing group
\1, \2      Backreference to group 1, 2
```

#### Alternation
```
abc|def     Match abc or def
(cat|dog)   Match cat or dog
```

### Advanced Features

#### Capture Groups in Replacement
```vbscript
Dim regex, date, formatted
Set regex = Server.CreateObject("G3REGEXP")

date = "2025-01-29"

' Rearrange date format
regex.Pattern = "(\d{4})-(\d{2})-(\d{2})"
formatted = regex.Replace(date, "$3/$2/$1")

Response.Write formatted  ' Output: 29/01/2025
```

#### Named Patterns
```vbscript
Dim regex
Set regex = Server.CreateObject("G3REGEXP")

' Define reusable patterns
Dim emailPattern
emailPattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"

regex.Pattern = emailPattern
```

### Performance Characteristics
- Pattern compilation happens once (cached)
- Fast matching with Go's regexp engine
- Global search efficient for large texts
- Minimal memory overhead
- Suitable for real-time processing

### Error Handling
- Invalid patterns return errors
- Empty matches handled gracefully
- Null/empty strings handled safely
- Error messages logged to server console

### Limitations
- Go regexp syntax (not PCRE)
- No lookahead/lookbehind in standard Go regexp
- Pattern size limited by memory
- No compiled pattern caching across calls

### VBScript RegExp Compatibility

**Supported Features**:
- Pattern, IgnoreCase, Global, Multiline properties
- Test, Execute, Replace methods
- Match objects with Index, Length, Value
- Matches collection with Count

**Not Supported**:
- Dynamic properties
- Compiled pattern persistence
- Replace callback functions

### Common Use Cases
1. Email and URL validation
2. Phone number formatting
3. Data extraction and parsing
4. Text replacement and cleanup
5. Form field validation
6. Log file analysis
7. HTML/XML tag removal
8. Password strength validation
9. Date format conversion
10. CSV/TSV processing

### Future Enhancements
- Lookahead/lookbehind support
- Compiled pattern caching
- Callback replacements
- Unicode property classes
- Performance optimizations
- Extended debugging information
