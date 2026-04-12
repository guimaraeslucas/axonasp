# Use G3TestSuite in AxonASP

## Overview
G3TestSuite is a native assertion object used by the axonasp-testsuite runner to execute Classic ASP unit tests inside the AxonASP VM.

## Syntax
```asp
<%
Dim t
Set t = Server.CreateObject("G3TestSuite")
%>
```

## Parameters and Arguments
- ProgID (String, Required): Use `G3TestSuite` or `G3Test`.
- Member access (Optional): Use the documented assertion methods and suite counters.

## Return Values
Returns a native object handle that records assertion counters and failure messages for the current ASP file.

## Remarks
- The object is optimized for backend test execution and keeps only compact counters plus failure strings.
- Member names are case-insensitive.
- Equality assertions reuse the VM comparison path so VBScript coercion stays aligned with runtime semantics.
- AssertRaises executes one Execute-compatible VBScript statement block and passes when that block raises a compile or runtime error.
- The testsuite runner aggregates every G3Test object created during one file execution.
- Example suites are available under `www/tests/testsuite/`.

## Code Example
```asp
<%
Dim t
Set t = Server.CreateObject("G3TestSuite")

t.Describe "String helpers"
t.AssertEqual "ABC", UCase("abc"), "UCase should normalize text"
t.AssertNotEqual "ABC", LCase("abc"), "LCase should produce a different value"
t.AssertTrue Len("abc") = 3, "Len should return 3"
t.AssertFalse IsEmpty("abc"), "String should not be Empty"
t.AssertNothing Nothing, "Nothing should pass AssertNothing"
t.AssertTypeName "String", "abc", "TypeName should be String"
t.AssertLength 3, Array("a", "b", "c"), "Array length should be 3"
t.AssertRaises "Err.Raise 13, ""docs"", ""type mismatch""", 13, "Err.Raise should surface explicit numbers"
%>
```