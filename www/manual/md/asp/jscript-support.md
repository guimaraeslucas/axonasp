# Use JScript in Classic ASP Pages

## Overview
This page explains how to run **JScript** in G3Pix AxonASP pages, how to declare the JScript language directive, how to call ASP intrinsic objects, and how to write runtime diagnostics with the global `console` object.

## Syntax
```asp
<%@ Language=JScript %>
<%
// JScript server-side code
Response.Write("Hello from JScript");
%>
```

## Parameters and Arguments
- **Language directive value** (required): must be `JScript` for the page block you want to execute as JScript.
- **ASP intrinsic object names** (optional in code, required for object access): `Request`, `Response`, `Server`, `Session`, `Application`, `Err`, and `ObjectContext`.
- **console methods** (optional): `console.log(value)`, `console.info(value)`, `console.warn(value)`, `console.error(value)`.
- **console method argument** (required): at least one argument. If the argument is an array or object, AxonASP serializes it to JSON before printing.

## Return Values
- The language directive returns no runtime value.
- ASP intrinsic object method calls return the value documented by each object member.
- `console.log`, `console.info`, `console.warn`, and `console.error` return no page output value. They write to console streams and optional log files.

## Remarks
- Use `Language=JScript` at the top of the ASP page to ensure server-side parsing as JScript.
- JScript can call ASP intrinsic objects directly, for example `Response.Write(...)` and `Request.QueryString("id")`.
- The global `console` object is available without object instantiation.
- `console.log` and `console.info` write to standard output.
- `console.warn` and `console.error` write to standard error.
- If `global.enable_log_files = true`, AxonASP appends `console.log` and `console.info` entries to `./temp/console.log`, and appends `console.warn` and `console.error` entries to `./temp/error.log`.
- The log files store timestamp, level, and message text. Decorative console symbols are not persisted in file output.

## Code Example
```asp
<%@ Language=JScript %>
<%
var userId;
userId = Request.QueryString("id");

if (!userId || userId === "") {
    userId = "anonymous";
}

Response.Write("User: " + userId + "<br>");

console.info("JScript request started");
console.log({ user: userId, source: "jscript-page" });
console.warn("Sample warning from JScript page");
%>
```
