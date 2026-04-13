# G3CRYPTO Properties

## Overview
This page lists the properties available in the **G3CRYPTO** library for managing cryptographic state and configuration.

## Property List

- **BCryptCost**: Read/Write. Configures the computational cost for bcrypt operations.
- **CanReuseTransform**: Read-only. Indicates if the transform can be reused (always True).
- **Hash**: Read-only. Returns the raw byte array of the most recent hash operation.
- **HashSize**: Read-only. Returns the bit-size of the most recent hash operation.

## Remarks
- Accessing properties is efficient and does not trigger expensive cryptographic operations.
- Use the **Hash** property when binary output is required for database storage or protocol transmission.
