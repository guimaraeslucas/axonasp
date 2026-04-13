# The #include Directive

## Overview
The **#include** directive is a server-side include (SSI) mechanism used in G3Pix AxonASP to insert the content of one file into another before the script is executed by the Virtual Machine. This directive is essential for code reuse, allowing developers to maintain shared components such as headers, footers, database connection strings, and library functions in a single location.

## Syntax
The directive must be placed inside HTML comments and uses either the **file** or **virtual** attribute:

```asp
<!--#include file="filename"-->
<!--#include virtual="/path/filename"-->
```

## Parameters and Arguments
- **file** (String): Specifies a path relative to the directory of the current file. You cannot use absolute paths or ".." to move up in the directory structure with this attribute.
- **virtual** (String): Specifies a path relative to the web root of the AxonASP application. This is the preferred method for including files located in global or shared directories.

## How it Works
The AxonASP Lexer identifies the **#include** directive during the initial parsing phase. The engine then retrieves the content of the specified file and physically injects it into the source stream, replacing the directive tag.

- **Preprocessing**: Includes are processed *before* any VBScript code is executed. You cannot use VBScript variables to dynamically set the path of an include directive.
- **Scope**: Variables, procedures, and classes defined in an included file are available to the parent script as if they were written directly within it.
- **Nesting**: Included files can themselves contain **#include** directives, allowing for deeply nested component architectures.
- **Circular References**: AxonASP prevents infinite recursion by limiting include depth and detecting circular references.

## Remarks
- **File Extensions**: While `.inc` and `.asp` are the most common extensions for included files, AxonASP processes any file type specified in the directive as script content.
- **Performance**: Excessive use of deeply nested includes can impact initial compilation time, though AxonASP's high-performance script cache mitigates this for subsequent requests.
- **Alternative**: For dynamic file inclusion during runtime, use `Server.Execute` or `Server.Transfer`.

## Code Example
The following example demonstrates a standard page structure using include directives for a consistent layout.

```asp
<!--#include virtual="/includes/config.inc"-->
<!--#include file="header.inc"-->

<%
' Main page logic
Response.Write "<h1>Welcome to the Dashboard</h1>"
Response.Write "<p>Connected to: " & DB_NAME & "</p>"
%>

<!--#include file="footer.inc"-->
```
