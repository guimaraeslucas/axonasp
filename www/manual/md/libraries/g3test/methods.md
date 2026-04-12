# Methods

## Overview

This page lists methods exposed by G3TestSuite.

## Method List
- Describe
- AssertEqual
- AssertNotEqual
- AssertTrue
- AssertFalse
- AssertEmpty
- AssertNull
- AssertNothing
- AssertTypeName
- AssertLength
- AssertCount
- AssertRaises
- Fail

## Remarks

Use one `Describe` call per logical block to improve failure output in axonasp-testsuite reports.

## Signatures
- `AssertEqual expected, actual, message`
- `AssertNotEqual expected, actual, message`
- `AssertTrue condition, message`
- `AssertFalse condition, message`
- `AssertEmpty value, message`
- `AssertNull value, message`
- `AssertNothing value, message`
- `AssertTypeName expectedType, value, message`
- `AssertLength expectedLength, value, message`
- `AssertCount expectedCount, value, message` (`AssertLength` alias)
- `AssertRaises code, message`
- `AssertRaises code, expectedNumber, message`
- `AssertRaises code, expectedText, message`

## Example
```asp
<%
Dim t
Dim emptyValue
Dim nullValue
Set t = Server.CreateObject("G3TestSuite")
nullValue = Null

t.Describe "Math"
t.AssertEqual 4, 2 + 2, "2 + 2 should equal 4"
t.AssertNotEqual 5, 2 + 2, "2 + 2 should not equal 5"
t.AssertTrue 5 > 1, "5 should be greater than 1"
t.AssertEmpty emptyValue, "Uninitialized values should be Empty"
t.AssertNull nullValue, "Explicit Null values should be Null"
t.AssertNothing Nothing, "Nothing should pass AssertNothing"
t.AssertTypeName "String", "abc", "TypeName(String) should match"
t.AssertLength 3, Array("a", "b", "c"), "Array length should be 3"
t.AssertCount 5, "Axon!", "String length should be 5"
t.AssertRaises "Err.Raise 13, ""docs"", ""type mismatch""", 13, "Err.Raise should preserve explicit numbers"
%>
```