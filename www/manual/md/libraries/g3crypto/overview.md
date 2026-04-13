# Use the G3CRYPTO Library

## Overview
The **G3CRYPTO** library provides high-performance cryptographic services for G3Pix AxonASP applications. It includes support for secure password hashing with bcrypt, industry-standard message digests (MD5, SHA-1, SHA-2, SHA-3), random data generation, and HMAC/PBKDF2 key derivation. The library is designed for zero-allocation performance and follows strict security standards for backend operations.

## Syntax
To instantiate the library, use the following syntax:
```asp
Set crypto = Server.CreateObject("G3CRYPTO")
```

## Prerequisites
No external dependencies are required. The G3CRYPTO library is a built-in native component of the AxonASP environment.

## How it Works
The G3CRYPTO object operates as a stateless service for most hashing operations, but it maintains internal state for the **Hash** property and **BCryptCost** configuration. When a hashing method (such as **SHA256**) is called, the library processes the input and stores the resulting byte array in an internal buffer, which can then be accessed via the **Hash** property if the raw binary output is needed.

For password security, the library uses the **bcrypt** algorithm, which includes automatic salting and configurable computational cost to protect against brute-force attacks.

## API Reference

### Methods
- **Blake2b256**: Computes a 256-bit BLAKE2b hash.
- **Blake2b512**: Computes a 512-bit BLAKE2b hash.
- **ComputeHash**: Computes a hash using a specified algorithm.
- **GetBCryptCost**: Retrieves the current bcrypt work factor.
- **HashPassword**: Creates a secure bcrypt hash for a password.
- **HmacSHA256**: Computes an HMAC using the SHA-256 algorithm.
- **HmacSHA512**: Computes an HMAC using the SHA-512 algorithm.
- **Initialize**: Resets the internal state and clears sensitive buffers.
- **MD5**: Computes a 128-bit MD5 hash.
- **PBKDF2SHA256**: Derives a key using PBKDF2-HMAC-SHA256.
- **RandomBase64**: Generates a cryptographically secure random string in Base64 format.
- **RandomBytes**: Generates an array of cryptographically secure random bytes.
- **RandomHex**: Generates a cryptographically secure random string in Hexadecimal format.
- **SetBCryptCost**: Configures the bcrypt work factor (cost).
- **SHA1**: Computes a 160-bit SHA-1 hash.
- **SHA256**: Computes a 256-bit SHA-256 hash.
- **SHA3_256**: Computes a 256-bit SHA-3 hash.
- **SHA3_512**: Computes a 512-bit SHA-3 hash.
- **SHA384**: Computes a 384-bit SHA-2 hash.
- **SHA512**: Computes a 512-bit SHA-2 hash.
- **UUID**: Generates a version 4 Universally Unique Identifier.
- **VerifyPassword**: Validates a plain-text password against a bcrypt hash.

### Properties
- **BCryptCost**: Gets or sets the work factor for bcrypt operations.
- **CanReuseTransform**: Indicates if the cryptographic transform can be reused.
- **Hash**: Returns the byte array of the last computed hash.
- **HashSize**: Returns the size of the last computed hash in bits.

## Code Example
The following example demonstrates how to hash a password and generate a secure random token.

```asp
<%
Dim crypto, password, hash, token
Set crypto = Server.CreateObject("G3CRYPTO")

' Securely hash a password
password = "UserSecret123!"
hash = crypto.HashPassword(password)
Response.Write "Password Hash: " & hash & "<br>"

' Verify the password
If crypto.VerifyPassword(password, hash) Then
    Response.Write "Verification Successful<br>"
End If

' Generate a random session token
token = crypto.RandomHex(32)
Response.Write "Random Token: " & token & "<br>"

Set crypto = Nothing
%>
```
