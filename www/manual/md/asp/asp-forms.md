# Handle ASP Form Data

## Overview

Use `Request.Form` to read fields submitted in an HTTP `POST` request in G3Pix AxonASP. Each key matches an HTML input `name` attribute.

`Request.Form` is built in. Do not instantiate it with `Server.CreateObject`.

## Prerequisites

- The page must receive data using HTTP `POST`.
- The submitted fields must include `name` attributes.
- Do not call `Request.BinaryRead` in the same request where you use `Request.Form`.

## Read a Single Field

### Syntax

```asp
value = Request.Form("fieldName")
```

### Return value

Returns a String containing the first submitted value for `fieldName`. Returns an empty String when `fieldName` is not present.

### Example

```asp
<%
Dim userName
userName = Trim(Request.Form("username"))

If userName = "" Then
    Response.Write "Username is required."
    Response.End
End If

Response.Write "Welcome, " & Server.HTMLEncode(userName)
%>
```

## Read Multi-Value Fields

### Syntax

```asp
value = Request.Form("fieldName")(index)
count = Request.Form("fieldName").Count
```

### Return value

- `Request.Form("fieldName")(index)` returns a String containing the value at the 1-based `index`.
- `Request.Form("fieldName").Count` returns an Integer with the number of values for that field.

### Example

```asp
<%
Dim i, selectedTagCount
selectedTagCount = Request.Form("tags").Count

For i = 1 To selectedTagCount
    Response.Write "Tag " & i & ": " & Server.HTMLEncode(Request.Form("tags")(i)) & "<br>"
Next
%>
```

## Enumerate Posted Keys

### Syntax

```asp
For Each key In Request.Form
    ' ...
Next
```

### Return value

`For Each` returns each key name as a String.

### Example

```asp
<%
Dim key
For Each key In Request.Form
    Response.Write key & " = " & Server.HTMLEncode(Request.Form(key)) & "<br>"
Next
%>
```

## How It Works

- The form collection is parsed from the request body when first accessed.
- Keys are case-insensitive.
- If you call `Request.BinaryRead`, AxonASP treats the body as raw bytes and does not provide `Request.Form` for that request.

## API Reference

| Member | Syntax | Returns |
|---|---|---|
| Field value | `Request.Form("fieldName")` | String (first value for field, or empty String if missing) |
| Indexed value | `Request.Form("fieldName")(index)` | String (value at 1-based index) |
| Value count | `Request.Form("fieldName").Count` | Integer (number of values for field) |
| Key enumeration | `For Each key In Request.Form` | String key per iteration |
