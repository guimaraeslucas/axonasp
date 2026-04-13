# G3CRYPTO Methods

## Overview
This page provides a summary of the methods available in the **G3CRYPTO** library. Each method is designed for high-performance cryptographic operations within the AxonASP environment.

## Method List

- **Blake2b256**: Computes a 256-bit BLAKE2b hash for high-speed data integrity.
- **Blake2b512**: Computes a 512-bit BLAKE2b hash for high-speed data integrity.
- **ComputeHash**: Computes a hash using a specified algorithm name.
- **GetBCryptCost**: Returns the current work factor for bcrypt hashing.
- **HashPassword**: Generates a secure bcrypt hash for password storage.
- **HmacSHA256**: Computes an HMAC using SHA-256 for message authentication.
- **HmacSHA512**: Computes an HMAC using SHA-512 for message authentication.
- **Initialize**: Resets the object state and clears the internal hash buffers.
- **MD5**: Computes a 128-bit MD5 message digest.
- **PBKDF2SHA256**: Derives a cryptographic key using the PBKDF2-HMAC-SHA256 algorithm.
- **RandomBase64**: Generates cryptographically secure random bytes in Base64 format.
- **RandomBytes**: Generates cryptographically secure random bytes as a VBScript array.
- **RandomHex**: Generates cryptographically secure random bytes in hexadecimal format.
- **SetBCryptCost**: Configures the work factor for the bcrypt hashing algorithm.
- **SHA1**: Computes a 160-bit SHA-1 message digest.
- **SHA256**: Computes a 256-bit SHA-256 message digest.
- **SHA3_256**: Computes a 256-bit SHA-3 (Keccak) message digest.
- **SHA3_512**: Computes a 512-bit SHA-3 (Keccak) message digest.
- **SHA384**: Computes a 384-bit SHA-2 message digest.
- **SHA512**: Computes a 512-bit SHA-2 message digest.
- **UUID**: Generates a version 4 Universally Unique Identifier.
- **VerifyPassword**: Validates a password against a previously generated bcrypt hash.

## Remarks
- Method names are case-insensitive.
- All hashing methods update the internal **Hash** and **HashSize** properties.
- Password-related methods (HashPassword/VerifyPassword) use the industry-standard bcrypt algorithm.
