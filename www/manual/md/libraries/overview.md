# Use G3Pix AxonASP Library Objects

## Overview
Use this page as a quick reference for object creation through `Server.CreateObject` in G3Pix AxonASP. It summarizes supported primary ProgIDs and common compatibility aliases.

## Prerequisites

- Run code inside a Classic ASP page hosted by G3Pix AxonASP.
- Use `Server.CreateObject("ProgID")` with a supported ProgID string.

## Syntax

```asp
Set obj = Server.CreateObject("ProgID")
```

## Object Creation Summary

| Category | Primary ProgID | Compatibility Aliases |
|---|---|---|
| Core functions | `G3AXON.FUNCTIONS` | `G3AXON` |
| Markdown | `G3MD` | None |
| String builder | `G3STRINGBUILDER` | None |
| Crypto | `G3CRYPTO` | None |
| JSON | `G3JSON` | None |
| Database helper | `G3DB` | None |
| HTTP helper | `G3HTTP` | None |
| Mail | `G3MAIL` | `CDONTS.NewMail`, `CDO.Message`, `Persits.MailSender` |
| Image | `G3IMAGE` | None |
| File helper | `G3FILES` | None |
| Template | `G3TEMPLATE` | None |
| Archive and compression | `G3ZIP`, `G3ZLIB`, `G3TAR`, `G3ZSTD`, `G3FC` | None |
| PDF | `G3PDF` | None |
| Upload | `G3FILEUPLOADER` | None |
| Shell | `WSCRIPT.SHELL` | `Shell` |
| ADOX | `ADOX.CATALOG` | None |
| ADODB | `ADODB.CONNECTION`, `ADODB.RECORDSET`, `ADODB.COMMAND`, `ADODB.STREAM`, `ADODBOLE.CONNECTION` | None |
| Scripting runtime | `SCRIPTING.FILESYSTEMOBJECT`, `SCRIPTING.DICTIONARY` | None |
| XML HTTP | `MSXML2.SERVERXMLHTTP` | `MSXML2.XMLHTTP`, `MICROSOFT.XMLHTTP` |
| XML DOM | `MSXML2.DOMDOCUMENT` | `MICROSOFT.XMLDOM` |
| MSWC compatibility | `MSWC.ADROTATOR`, `MSWC.BROWSERTYPE`, `MSWC.NEXTLINK`, `MSWC.CONTENTROTATOR`, `MSWC.COUNTERS`, `MSWC.PAGECOUNTER`, `MSWC.TOOLS`, `MSWC.MYINFO`, `MSWC.PERMISSIONCHECKER` | None |
| RegExp | `VBSCRIPT.REGEXP` | `REGEXP` |

## Return Value

`Server.CreateObject` returns an object handle bound to the native G3Pix AxonASP implementation for the resolved ProgID.

## Remarks

- **Object creation is case-insensitive for supported ProgIDs.**
- **Unsupported ProgIDs fall through to host-level CreateObject handling.**
- **Compatibility aliases are mapped to native AxonASP implementations where available.**

## Code Example

```asp
<%
Option Explicit
Dim ax, db, stm, xmlHttp

Set ax = Server.CreateObject("G3AXON.FUNCTIONS")
Set db = Server.CreateObject("ADODB.Connection")
Set stm = Server.CreateObject("ADODB.Stream")
Set xmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")

Response.Write TypeName(ax)

Set xmlHttp = Nothing
Set stm = Nothing
Set db = Nothing
Set ax = Nothing
%>
```

## API Reference

- Entry point: `Server.CreateObject`
- Input: ProgID string
- Resolution: native AxonASP ProgID map, then compatibility alias map, then host-level fallback
- Output: object handle for the resolved implementation
