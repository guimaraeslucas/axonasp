CRITICAL TASK:
You are tasked with generating and expanding the documentation for the G3Crypto, G3Files, G3FileUploader, G3Http native functions, which are instantiated in Classic ASP via Server.CreateObject("[LibraryName]").

You must read the existing library located at axonvm/lib_g3crypto.go, axonvm/lib_g3files.go, axonvm/lib_g3fileuploader.go and generate the heavily expanded, finalized output into mcp/docs.md.

ABSOLUTE FORMATTING CONSTRAINTS (DO NOT IGNORE):
This Markdown file (mcp/docs.md) is ingested by a Go-based parser that uses fuzzy string matching to serve this data to an MCP (Model Context Protocol) server for AI coding assistants.

DO NOT create new fields.

DO NOT modify, translate, or rename the existing field labels.

DO NOT change the markdown heading levels or bolding structure.

MUST strictly follow the exact template provided below for every single function. Any deviation will break the parser.

DO NOT change markdown already present in the existing documentation. You may only expand the content within the fields, not the structure or formatting. 

DO NOT include any additional commentary, explanations, or notes outside of the specified fields. The output must be strictly limited to the content within the fields as per the template.

*DO NOT change the existing content in the file*, as it is already populated with some functions. You are only to expand the content of the existing functions and add new functions following the same format.

REQUIRED TEMPLATE:

Markdown
## ServerObject: [LibraryName] Function: [Function Name(parameters)]
**Keywords:** [Comma-separated list]
**Description:** [Detailed explanation]
**Observations:** [Technical notes]
**Syntax:**
```vbscript
[Code Example]
```
(Note: Do not include the brackets [], they are placeholders, in empty parameters functions leave just ()).

CONTENT QUALITY GUIDELINES (AI-Focused):
Since this documentation is meant to be read by another AI Coding Agent to help it write code, the content must be highly technical, precise, and unambiguous.

Keywords: Provide a rich, comma-separated list of synonyms, actions, and use-cases. Think about what an AI would search for to find this function (e.g., if it's a date formatting tool, use: date, format, time, parse, string manipulation, conversion).

Description: Explain exactly what the function does. You MUST include:

The expected data types for all input parameters.

The expected data type of the return value.

How the data should be handled or manipulated.

Observations: Detail any edge cases, platform-specific behaviors (e.g., Windows vs. Linux differences in the Go backend), limitations, performance considerations, or specific properties the developer/AI must be aware of. Remember to write VALID Asp Classic/VBScript examples and syntax. Don't forget the way If Then are implemented and how they work in ASP Classic (remember the new line requirement) and to use End If.

Syntax: Provide a clear, working VBScript example. Every example MUST start with the object instantiation:

Dim ax
Set ax = Server.CreateObject("[LibraryName]")
' Followed by the function call

Execution Steps:
Read the files instructed above to identify all the functions defined in the Go library that are exposed to the user via Server.CreateObject(). For each function, extract the necessary information to fill in the fields in the template (Keywords, Description, Observations, Syntax).
For each function found, expand the description and examples significantly to meet the AI-Focused quality guidelines above.
Only expose functions that are defined in the Go library and available to the user. Do not create hypothetical functions or include any functions that have been removed from the Go code or are just used internally.
Append the strictly formatted output content to mcp/docs.md.