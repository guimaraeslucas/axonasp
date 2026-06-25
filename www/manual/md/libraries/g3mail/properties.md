# G3MAIL Properties

## Overview

This page lists the properties exposed by `G3MAIL`.

## Properties

| Property | Access | Type | Description |
|---|---|---|---|
| `Host` | Read/Write | String | SMTP host name. |
| `Port` | Read/Write | Integer | SMTP port number. |
| `Username` | Read/Write | String | SMTP authentication username. |
| `Password` | Read/Write | String | SMTP authentication password. |
| `From` | Read/Write | String | Sender email address used in the message header. |
| `FromName` | Read/Write | String | Display name for the sender. |
| `To` | Read/Write | String | Recipient list for To, represented as comma-separated text. |
| `CC` | Read/Write | String | Recipient list for CC, represented as comma-separated text. |
| `BCC` | Read/Write | String | Recipient list for BCC, represented as comma-separated text. |
| `Subject` | Read/Write | String | Message subject line. |
| `Body` | Read/Write | String | Message body text. |
| `HTMLBody` | Read/Write | String | Message HTML body. |
| `IsHTML` | Read/Write | Boolean | Toggles HTML (`True`) or plain-text (`False`) body mode. |
| `BodyFormat` | Read/Write | Integer | Body format selector (`0` for HTML, `1` for plain text). |

## Remarks

- Instantiate the library with `Server.CreateObject("G3MAIL")`.
- Property names are case-insensitive.
