# G3TESTSUITE Properties

## Overview
This page provides a summary of the properties available in the **G3TESTSUITE** library for inspecting the state and result of a test session.

## Property List

- **CurrentDescribe**: Read-only. Returns a **String** containing the active description label.
- **Failed**: Read-only. Returns an **Integer** representing the number of failed assertions.
- **HasFailures**: Read-only. Returns a **Boolean** indicating if any failed assertions have been recorded.
- **Passed**: Read-only. Returns an **Integer** representing the number of successful assertions.
- **Suite**: Read/Write. Gets or sets a **String** for the current test block description.
- **Total**: Read-only. Returns an **Integer** representing the total count of assertions executed.

## Remarks
- Properties like **Failed** and **Passed** are automatically updated upon completion of every assertion method call.
- Use the **HasFailures** property as a quick flag to determine if the current test file has encountered errors.
- **CurrentDescribe** and **Suite** are identical and provide access to the label set by the **Describe** method.
