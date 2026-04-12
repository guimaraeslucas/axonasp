# Write New AxonASP Manual Pages

## Overview
Use this procedure when creating or updating AxonASP manual pages. The goal is consistent structure, accurate API coverage, and compatibility with tree-view rendering.

## Syntax
```text
# <Action-Oriented Title>

## Overview
## Syntax
## Parameters and Arguments
## Return Values
## Remarks
## Code Example
```

## Parameters and Arguments
- Action-Oriented Title (Required): Include object/member purpose directly in the title.
- Overview (Required): Short explanation of what the page covers.
- Syntax (Required): Exact Classic ASP/VBScript syntax form.
- Parameters and Arguments (Required): Include data type and required/optional state.
- Return Values (Required): Describe expected output and failure behavior.
- Remarks (Required): Include compatibility notes, edge cases, and performance details.
- Code Example (Required): Complete runnable Classic ASP example.

## Return Values
Following this template produces a manual page that is consistent with AxonASP standards and suitable for both users and coding agents.

## Remarks
- Do not create markdown hyperlinks inside manual content pages.
- In www/manual/menu.md, always use markdown links for navigable entries with the exact format &lsqb;Label&rsqb;&lpar;folder/document.md&rpar;.
- Keep folder/group labels in menu.md as plain bullet items when they must have child items.
- Menu link targets must be paths relative to www/manual/md.
- Keep all text and examples in English (US).
- Validate member names against axonvm/lib_*.go dispatch surfaces before publishing.
- Prefer predictable names such as overview.md, methods.md, properties.md, and family-specific pages where needed.
- For AI-generated ASP examples, enforce ASP_CODE_GUIDELINES.MD through the authoring/llm-classic-asp-coding.md checklist.

## Code Example
```asp
<%
Dim db
Set db = Server.CreateObject("G3DB")
db.Driver = "sqlite"
db.DSN = "./database.db"
If db.Open() Then
  Response.Write "Connected"
  db.Close
End If
%>
```
