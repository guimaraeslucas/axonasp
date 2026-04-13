# Write ASP Script Syntax

## Overview

Classic ASP pages in G3Pix AxonASP combine HTML markup with server-side VBScript. Code inside `<%` and `%>` runs on the server before the response is sent to the client.

Use this page as a baseline syntax reference for script blocks, output, directives, comments, and control flow.

## Prerequisites

- Save the page with `.asp` extension.
- Keep directives at the top of the file.
- Use VBScript syntax consistently in server blocks.

## Core Delimiters

### Execute server-side code

```asp
<%
Dim message
message = "Hello from server code"
Response.Write message
%>
```

### Output an expression directly

```asp
<p>Current date: <%= Date() %></p>
```

### Return value

- `<% ... %>` executes statements and returns no direct inline value.
- `<%= expression %>` evaluates `expression` and writes its String representation to the response output stream.

## Write Output with Response.Write

### Syntax

```asp
Response.Write expression
```

### Return value

Returns no value. The method appends the evaluated text to the HTTP response body.

### Example

```asp
<%
Response.Write "<h1>Dashboard</h1>"
Response.Write "<p>Welcome, " & Server.HTMLEncode("alice") & "</p>"
%>
```

## Use Page Directives

Place directives at the top of the file.

```asp
<%@Language="VBSCRIPT" CodePage="65001"%>
<%
Response.CodePage = 65001
Response.Charset = "utf-8"
%>
```

### Return value

Directives do not return values. They configure how AxonASP parses and executes the page.

## Add Comments

```asp
<%
' Single-line VBScript comment
Response.Write "Comments are ignored by the runtime"
%>
```

### Return value

Comments do not return values and do not produce output.

## Basic Control Flow

### If...Then...Else

```asp
<%
Dim isAdmin
isAdmin = True

If isAdmin Then
    Response.Write "<p>Admin panel enabled.</p>"
Else
    Response.Write "<p>Standard user mode.</p>"
End If
%>
```

### For...Next

```asp
<%
Dim i
For i = 1 To 3
    Response.Write "Item " & i & "<br>"
Next
%>
```

### Return value

Control statements do not return values. They only control execution order.

## How It Works

- AxonASP parses HTML and server blocks in one page request.
- Server-side VBScript runs first.
- The resulting output is combined with static HTML and sent to the client.

## API Reference

| Syntax element | Form | Returns |
|---|---|---|
| Script block | `<% statements %>` | No direct inline value |
| Expression block | `<%= expression %>` | Writes String representation of expression |
| Output method | `Response.Write expression` | No return value |
| Directive block | `<%@ ... %>` | No return value |
| Comment | `' comment` | No return value |
| Conditional | `If ... Then ... End If` | No return value |
| Loop | `For ... Next` | No return value |
