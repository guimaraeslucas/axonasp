# Use axonasp-testsuite

## Overview

`axonasp-testsuite` is the native Classic ASP test runner for AxonASP. Conceptually, it serves the same role as PHPUnit for PHP projects or `go test` for Go projects: discover tests, execute them, aggregate assertions, print a human-friendly summary, and return a process exit code suitable for CI/CD pipelines.

The runner executes test pages through the AxonASP VM, collects assertion reports from each `G3Test` or `G3TestSuite` object created in each file, and fails the process when any suite has failed assertions, runtime errors, compile errors, or no assertions.

## Discovery Rules

The runner recursively scans a directory and executes files that match:

- `*test.asp`
- `test_*.asp`

The default extension list comes from `global.execute_as_asp` in `config/axonasp.toml`. Files are sorted before execution to provide deterministic run order.

## Command-Line Usage

Windows:

```powershell
.\axonasp-testsuite.exe .\www\tests\testsuite
```

Linux and macOS:

```bash
./axonasp-testsuite ./www/tests/testsuite
```

You can run a broader folder:

```powershell
.\axonasp-testsuite.exe .\www\tests
```

## How Test Files Are Executed

For each discovered file, the runner performs this flow:

1. Loads or compiles the ASP script through the bytecode cache.
2. Creates a VM instance from the cached program.
3. Sets a CLI-style ASP host context (`REQUEST_METHOD=CLI`) with server root and request path.
4. Executes the file and captures output and assertion reports.
5. Prints per-file status and aggregate summary.

A file is marked as failed when any of the following happens:

- Compile error
- Runtime error
- One or more failed assertions
- No assertions executed (no `G3Test`/`G3TestSuite` report generated)

## Creating ASP Tests

Inside each test file, instantiate the test object with:

```asp
Set t = Server.CreateObject("G3TestSuite")
```

`"G3Test"` is also accepted as an alias:

```asp
Set t = Server.CreateObject("G3Test")
```

### Minimal Smoke Test

```asp
<%
Option Explicit

Dim t
Set t = Server.CreateObject("G3TestSuite")

t.Describe "G3TestSuite smoke test"
t.AssertEqual 4, 2 + 2, "2 + 2 should equal 4"
t.AssertTrue Len("axon") = 4, "Len should report the expected size"
t.AssertFalse IsEmpty("ready"), "Literal text should not be Empty"
%>
```

## Assertion API

The following methods are supported by the native `G3Test` object.

### Describe

Sets the active suite/block label for failure messages.

```asp
t.Describe "Math operations"
```

### AssertEqual(expected, actual, [message])

Passes when values are equal under VM comparison/coercion semantics.

### AssertNotEqual(expected, actual, [message])

Passes when values are different.

### AssertTrue(condition, [message])

Passes when condition is truthy under VBScript boolean coercion.

### AssertFalse(condition, [message])

Passes when condition is falsy under VBScript boolean coercion.

### AssertEmpty(value, [message])

Passes when value is VBScript `Empty`.

### AssertNull(value, [message])

Passes when value is VBScript `Null`.

### AssertNothing(value, [message])

Passes when value is `Nothing`-compatible in AxonASP semantics.

### AssertTypeName(expectedType, value, [message])

Passes when `TypeName(value)` matches `expectedType` (case-insensitive).

### AssertLength(expectedLength, value, [message])

Passes when the target value length/count equals expected length.

Supported targets include:

- Strings (character length)
- Arrays (element count)
- `Empty` (treated as length 0)
- Native objects exposing `Count`

### AssertCount(expectedCount, value, [message])

Alias of `AssertLength`.

### AssertRaises(code, [expected], [message])

Executes dynamic VBScript code and verifies an error is raised.

Supported argument forms:

- `AssertRaises(code, message)`
- `AssertRaises(code, expectedNumber, message)`
- `AssertRaises(code, expectedText, message)`

Examples:

```asp
t.AssertRaises "Err.Raise 13, \"suite\", \"type mismatch\"", 13, "Should raise Err.Number 13"
t.AssertRaises "Err.Raise 13, \"suite\", \"type mismatch\"", "type mismatch", "Should contain error text"
t.AssertRaises "Function Broken(", "Syntax errors should be trapped"
```

### Fail([message])

Forces an explicit failure.

```asp
t.Fail "This branch should be unreachable"
```

## Suite Properties

These read properties are available:

- `Suite` / `Description` / `CurrentDescribe`
- `Total` / `TotalTests`
- `Passed`
- `Failed`
- `HasFailures`

The suite name can also be set through writable aliases:

```asp
t.Suite = "Authentication"
```

## Practical Test Organization

- Keep one feature area per file (`auth_test.asp`, `json_test.asp`).
- Start each block with `Describe` to make failure output easy to scan.
- Use explicit assertion messages that describe intent, not implementation.
- Prefer deterministic data; avoid dependence on clock/timezone unless explicitly testing those behaviors.
- Keep one assertion intent per line so failed output maps directly to a single behavior.

## Example: Multi-Assertion Test File

```asp
<%
Option Explicit

Dim t, arr
Set t = Server.CreateObject("G3TestSuite")

t.Describe "Core VBScript compatibility"
t.AssertEqual 10, CLng(9.6), "CLng should round using VBScript rules"
t.AssertTypeName "String", CStr(123), "CStr should produce String"
t.AssertNotEqual "abc", "xyz", "Different strings must not be equal"

arr = Array("a", "b", "c")
t.AssertCount 3, arr, "Array should expose element count"

t.Describe "Error behavior"
t.AssertRaises "Err.Raise 5, \"suite\", \"invalid procedure call\"", 5, "Err.Raise number should match"
%>
```

## CI Usage

Example CI step:

```yaml
- name: Build axonasp-testsuite
  run: go build -o axonasp-testsuite ./testsuite

- name: Run ASP test suite
  run: ./axonasp-testsuite ./www/tests/testsuite
```

Use process exit code:

- `0` when all suites pass
- `1` when any suite fails or when execution is invalid

## Runtime Integration Notes

- `global.asa` is loaded before suite execution and application/session end hooks are called when the run finishes.
- The runner uses the shared VM pool and script cache for fast execution.
- Output includes per-file PASS/FAIL status, assertion counters, and detailed failure messages.
- If a test file writes to output, that output is displayed under the file report for diagnostics.

## Remarks

- Treat `axonasp-testsuite` as your authoritative regression gate for ASP behavior changes.
- Prefer `G3TestSuite` in documentation and project examples; use `G3Test` alias only when needed.
- Ensure every test file executes at least one assertion; files with no assertions are treated as failure by design.