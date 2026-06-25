# G3MAIL Methods

## Overview

This page summarizes every method exposed by `G3MAIL` in G3Pix AxonASP.

## Methods

| Method | Returns | Description |
|---|---|---|
| `AddAddress(address)` | Boolean | Adds one recipient address to the To list. Returns `True`. |
| `AddCC(address)` | Boolean | Adds one recipient address to the CC list. Returns `True`. |
| `AddBCC(address)` | Boolean | Adds one recipient address to the BCC list. Returns `True`. |
| `Clear()` | Boolean | Clears To/CC/BCC/Subject/Body context and resets HTML mode. Returns `True`. |
| `AddAttachment(filepath)` | Boolean | Attaches a file to the email. Returns `True` on success. |
| `AddRelatedBodyPart(filepath, cid)` | Object | Embeds a related resource (e.g. an image) with the specified Content-ID (CID). Returns a body part object. |
| `Send()` | Boolean or String | Sends using configured properties and returns `True` on success; returns an error String when SMTP configuration is incomplete or send fails. |
| `Send(to, subject, body)` | Boolean or String | CDONTS-style overload that first sets To/Subject/Body, then sends. Returns `True` on success; returns an error String on failure. |

## Remarks

- Instantiate the library with `Server.CreateObject("G3MAIL")`.
- Method names are case-insensitive.
- `Send` does not return Empty for operational failure; it returns an error string.
