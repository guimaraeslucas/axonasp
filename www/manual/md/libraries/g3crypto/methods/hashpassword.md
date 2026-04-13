# HashPassword Method

## Overview

Hashes a password string using the bcrypt algorithm and the configured computational cost with the G3Pix AxonASP G3CRYPTO library.

## Syntax

```asp
hashedPassword = obj.HashPassword(password)
```

## Parameters

- **password** (String): The plain-text password to be hashed.

## Return Values

Returns a String containing the bcrypt-hashed password (including the algorithm identifier, cost, and salt).

## Remarks

- Instantiated via `Server.CreateObject("G3CRYPTO")`.
- The `BCryptCost` property determines the computational difficulty of the hash.
- Use the `VerifyPassword` method to check a plain-text password against a hashed value.

## Code Example

```asp
<%
Dim crypto, hashedPassword
Set crypto = Server.CreateObject("G3CRYPTO")
hashedPassword = crypto.HashPassword("mySecretPassword")
Response.Write "Secure password hash: " & hashedPassword
Set crypto = Nothing
%>
```
