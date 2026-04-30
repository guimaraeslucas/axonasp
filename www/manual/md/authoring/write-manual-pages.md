# Write New AxonASP Manual Pages

## Overview
This document provides the standard procedure and strict rules for creating or updating AxonASP manual pages. It is designed to be read by both human technical writers and Large Language Models (LLMs) acting as documentation generators. The goal is consistent structure, precise API coverage, strict adherence to legacy standards, and compatibility with our UI tree-view rendering.

**DO NOT write generic or assumed documentation.** Be extremely specific about AxonASP's implementation.

## Syntax Template
Every manual page must follow this structure when authoring content for a specific object, method, or property:
```text
# <Action-Oriented Title>

## Overview
## Syntax
## Parameters and Arguments
## Return Values
## Remarks
## Code Example
```

## Section Requirements
- **Action-Oriented Title (Required):** Include the object/member purpose directly in the title.
- **Overview (Required):** Short, precise explanation of what the page covers.
- **Syntax (Required):** Exact Classic ASP/VBScript syntax form.
- **Parameters and Arguments (Required):** Include data type, required/optional state, and specific expected values. No generic descriptions.
- **Return Values (Required):** Describe the exact expected output and failure behavior. Never use generic catch-all descriptions.
- **Remarks (Required):** Include compatibility notes, edge cases, performance details, and AxonASP-specific behaviors.
- **Code Example (Required):** Complete runnable Classic ASP example following the `llm-classic-asp-coding.md` guidelines.

## Rules & Constraints
Whether you are a human writer or an LLM, you must strictly follow these rules:

1. **Format & Style:**
   - Follow the classic Microsoft Writing Style Guide (action-oriented titles, brief overviews, prerequisites, code examples, extra information on how the code works, and complete API references).
   - Use active voice, simple language, scannable lists, and **bold text** for emphasis.
   - **ABSOLUTELY NO EMOJIS.**
   - **NO MARKDOWN LINKS inside the content page.** Navigation links are strictly reserved for `www/manual/menu.md`.
   - Keep all text and examples in English (US).

2. **Branding:**
   - Use the "G3Pix AxonASP" name and branding.
   - **DO NOT** use "MSDN" or "Microsoft" names, trademarks, or logos anywhere in the text. Replicate their classic, rigorous aesthetic and structural style only.

3. **Instantiation Precision (ProgID):**
   - Never describe an object creation using aliases in the code snippet.
   - *Incorrect:* `Set obj = Server.CreateObject("G3CRYPTO and aliases")`
   - *Correct:* `Set obj = Server.CreateObject("G3CRYPTO")` (Always use the primary name only).

4. **Explicit Return Values:**
   - Do NOT use generic catch-all descriptions for return values.
   - *Incorrect:* "Returns a Variant result. Depending on the operation, this can be String, Boolean, Number, Array, Dictionary/object handle, or Empty."
   - *Correct:* You must document exactly what *that specific function* returns (e.g., "Returns a String containing the SHA-256 hash of the input", or "Returns a Boolean indicating whether the decryption was successful").
   - **DON'T BE GENERALIST. BE VERY SPECIFIC.**

5. **Menu Navigation (`www/manual/menu.md`):**
   - Always use markdown links for navigable entries with the exact format `[Label](folder/document.md)`.
   - Keep folder/group labels in `menu.md` as plain bullet items when they must have child items.
   - Menu link targets must be paths relative to `www/manual/md`.

6. **Summary Pages (`methods.md` and `properties.md`):**
   - When creating summary or index pages for a library (e.g., `/<library_name>/methods.md` or `/<library_name>/properties.md`), you MUST use markdown tables to list the members.
   - For `methods.md`, use the columns: `Method`, `Returns`, `Description` (e.g., see `/g3crypto/methods.md` style).
   - For `properties.md`, use the columns: `Property`, `Access`, `Type`, `Description`.

## LLM Authoring Prompt
If you are using an LLM to generate or update documentation for AxonASP, provide the AI with the source code (e.g., the `lib_*.go` file) and use the following prompt to ensure strict compliance:

```text
Act as an expert technical writer for the AxonASP project (developed by G3Pix). We need to write or significantly improve the documentation for the provided AxonASP library/object. Your writing is focused on ASP Classic, GoLang and VBScript.

Your task is to write the documentation to be precise, professional, and strictly compliant with our documentation standards.

**STRICT RULES & CONSTRAINTS (DO NOT IGNORE):**

1. **Format & Style:** Follow the Microsoft Writing Style Guide (action-oriented titles, brief overviews, prerequisites, code examples, extra information on how the code works, and complete API references). Use active voice, simple language, scannable lists, and bold text for emphasis. ABSOLUTELY NO EMOJIS. NO MARKDOWN LINKS inside the content page.
2. **Template:** Every manual page must follow this structure when authoring content for a specific object, method, or property - *Use exactly these headers*: `# Title`, `## Overview`, `## Syntax`, `## Parameters and Arguments`, `## Return Values`, `## Remarks`, `## Code Example`. For summary pages (`methods.md` or `properties.md`), use markdown tables. `methods.md` needs `Method | Returns | Description`. `properties.md` needs `Property | Access | Type | Description`.
3. **Branding:** Use the "G3Pix AxonASP" name. DO NOT use "MSDN" or "Microsoft" names or logos.
4. **Instantiation Precision (ProgID):** Never describe an object creation using aliases. e.g., use `Set obj = Server.CreateObject("PRIMARY_NAME_ONLY")`.
5. **Explicit Return Values:** Do NOT use generic catch-all descriptions. Document exactly what THAT SPECIFIC function returns (e.g., "Returns a String containing..."). DON'T BE GENERALIST. BE VERY SPECIFIC.
6. **Code Example:** Provide a complete, runnable Classic ASP example in English (US). Ensure the code is realistic, correct and demonstrates actual usage.
7. When creating code you can also check llm-classic-asp-coding.md in the manual, to ensure the ASP Classic code is following the guidelines.

Here is the source code / library details to document:
[INSERT SOURCE CODE OR DETAILS HERE]
```

## Example Code Guideline
```asp
<%
Dim crypto
' Correct primary instantiation
Set crypto = Server.CreateObject("G3CRYPTO")

Dim hashValue
' Precise documentation of the return value is expected in the text above this example
hashValue = crypto.MD5("my_secret_string")

If Len(hashValue) > 0 Then
    Response.Write "Hash generated successfully."
End If

Set crypto = Nothing
%>
```