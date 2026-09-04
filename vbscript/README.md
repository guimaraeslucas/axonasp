# VBScript Lexer & Tokenizer (`vbscript`)

The `vbscript` package implements the lexical analysis and token stream generation for VBScript and Classic ASP templates in AxonASP.  This code was adapted from https://github.com/kmvi/vbscript-parser/ ensuring compatibility with VBScript language specifications.

## Features

- **Multi-Mode Lexer**: Supports `ModeVBScript` for pure script parsing and `ModeASP` for ASP delimiters (`<% %>`, `<%= %>`, `<%@ %>`), server-side script tags (`<script runat="server">`), and `#include` directives.
- **Fast Tokenization**: Tokenizes source code into structured tokens ready for direct consumption by the single-pass compiler in `axonvm`.
- **Microsoft VBScript Compliance**: Follows case-insensitivity, syntax rules, operator precedence, and standard error categorization (`vberrorcodes.go`).
