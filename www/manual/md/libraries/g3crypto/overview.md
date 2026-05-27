# Use the G3CRYPTO Library

## Overview

`G3CRYPTO` provides native cryptographic operations for G3Pix AxonASP applications. It supports message digests, HMAC signatures, password hashing with bcrypt, PBKDF2 key derivation, secure random generation, and UUID creation.

## Prerequisites

No external dependencies are required.

## Syntax

```asp
Dim crypto
Set crypto = Server.CreateObject("G3CRYPTO")
```
```javascript
var crypto = Server.CreateObject("G3CRYPTO");
```
## How it Works

The object computes digests either as hex strings (for methods such as `SHA256`) or as raw byte arrays (for `ComputeHash`).

The object stores state for:

- `Hash`: Last raw digest bytes produced by digest operations that update internal hash state.
- `BCryptCost`: Configurable work factor used by bcrypt password hashing.

Password operations use bcrypt. Key derivation uses PBKDF2-HMAC-SHA256. Random generators use cryptographically secure OS randomness.

## API Reference

### Object

- **ProgID**: `G3CRYPTO`

### Method Categories

- **Digest Methods**: `MD5`, `SHA1`, `SHA256`, `SHA384`, `SHA512`, `SHA3_256`, `SHA3_512`, `Blake2b256`, `Blake2b512`
- **Digest Bytes Method**: `ComputeHash`
- **Password Methods**: `HashPassword`, `VerifyPassword`, `SetBCryptCost`, `GetBCryptCost`
- **MAC/KDF Methods**: `HmacSha256`, `HmacSha512`, `Pbkdf2Sha256`
- **Random/Identifier Methods**: `RandomBytes`, `RandomHex`, `RandomBase64`, `UUID`
- **State Method**: `Initialize`

### Properties

- **BCryptCost** (Read/Write): bcrypt work factor.
- **CanReuseTransform** (Read-only): Always `True`.
- **Hash** (Read-only): Last raw digest bytes.
- **HashSize** (Read-only): Digest size in bits for the current context.

## Code Example

The following example shows password hashing, verification, and secure token generation.

```asp
<%
Option Explicit
Dim crypto, password, hash, token
Set crypto = Server.CreateObject("G3CRYPTO")

password = "UserSecret123!"
hash = crypto.HashPassword(password)
Response.Write "Password Hash: " & hash & "<br>"

If crypto.VerifyPassword(password, hash) Then
    Response.Write "Verification successful<br>"
End If

token = crypto.RandomHex(32)
Response.Write "Random Token: " & token & "<br>"

Set crypto = Nothing
%>
```
