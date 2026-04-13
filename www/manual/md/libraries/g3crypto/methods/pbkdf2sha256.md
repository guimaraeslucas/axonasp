# Pbkdf2Sha256 Method

## Overview

Derives a cryptographic key from a password and salt using the PBKDF2 (Password-Based Key Derivation Function 2) with HMAC-SHA256 in the G3Pix AxonASP G3CRYPTO library.

## Syntax

```asp
result = obj.Pbkdf2Sha256(password, salt, [iterations], [keyLength])
```

## Parameters

- **password** (String): The master password to derive the key from.
- **salt** (String): A random salt string used to ensure unique output for identical passwords.
- **iterations** (Integer, Optional): The number of hashing iterations. Defaults to 100,000.
- **keyLength** (Integer, Optional): The desired length of the derived key in bytes. Defaults to 32 (256 bits).

## Return Values

Returns a String containing the derived key encoded as a lowercase hexadecimal string.

## Remarks

- Instantiated via `Server.CreateObject("G3CRYPTO")`.
- PBKDF2 is designed to be computationally expensive to resist brute-force attacks on passwords.
- The resulting key is also stored in the `Hash` property as a byte array.

## Code Example

```asp
<%
Dim crypto, salt, derivedKey
Set crypto = Server.CreateObject("G3CRYPTO")
salt = "random_salt_value"
derivedKey = crypto.Pbkdf2Sha256("mySecretPassword", salt, 100000, 32)
Response.Write "Derived Key: " & derivedKey
Set crypto = Nothing
%>
```
