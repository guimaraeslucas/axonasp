# Use the G3TESTSUITE Library

## Overview
The **G3TESTSUITE** library provides a native assertion framework for G3Pix AxonASP. It is the primary tool for writing automated unit tests and integration tests within the AxonASP environment. The library enables developers to verify variable states, compare values using native VBScript coercion, and validate error-handling logic.

## Syntax
To instantiate the test suite, use the following syntax:
```asp
Set t = Server.CreateObject("G3TESTSUITE")
```

## Prerequisites
No external dependencies are required. The G3TESTSUITE library is a built-in component of the AxonASP Virtual Machine and is designed to work seamlessly with the `axonasp-testsuite` runner.

## How it Works
The G3TESTSUITE object maintains internal counters for total, passed, and failed assertions. When an assertion method (such as **AssertEqual**) is called, the library evaluates the condition using the same comparison logic as the AxonASP VM, ensuring that test behavior perfectly matches production behavior.

If an assertion fails, the library records a failure message, including the active description set by **Describe**. These failures are aggregated by the AxonASP test runner to produce structured reports.

## API Reference

### Methods
- **AssertEmpty**: Verifies that a value is the VBScript `Empty` variant.
- **AssertEqual**: Verifies that two values are equal using standard coercion.
- **AssertFalse**: Verifies that a condition evaluates to `False`.
- **AssertLength**: Verifies the number of elements in an array or the length of a string.
- **AssertNothing**: Verifies that a value is the VBScript `Nothing` object.
- **AssertNotEqual**: Verifies that two values are not equal.
- **AssertNull**: Verifies that a value is the VBScript `Null` variant.
- **AssertRaises**: Verifies that a VBScript code block raises a specific error.
- **AssertTrue**: Verifies that a condition evaluates to `True`.
- **AssertTypeName**: Verifies the VBScript `TypeName` of a value.
- **Describe**: Sets a label for the current block of assertions.
- **Fail**: Explicitly records an assertion failure.

### Properties
- **Failed**: Returns the number of failed assertions in the current instance.
- **HasFailures**: Returns a Boolean indicating if any assertions have failed.
- **Passed**: Returns the number of passed assertions in the current instance.
- **Suite**: Gets or sets the current description label.
- **Total**: Returns the total number of assertions executed.

## Code Example
The following example demonstrates how to perform a series of assertions to validate a business logic function.

```asp
<%
Dim t, result
Set t = Server.CreateObject("G3TESTSUITE")

t.Describe "Financial Calculator - Interest Rates"

' Validate simple calculation
result = CalculateInterest(1000, 0.05, 1)
t.AssertEqual 50, result, "Interest on 1000 at 5% for 1 year should be 50"

' Validate type safety
t.AssertTypeName "Double", result, "Result must be a Double type"

' Validate error handling for negative values
t.AssertRaises "CalculateInterest -100, 0.05, 1", 5, "Should raise Illegal Function Call"

Response.Write "Tests completed. Passed: " & t.Passed & "/" & t.Total
Set t = Nothing
%>
```
