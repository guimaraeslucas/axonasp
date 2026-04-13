# UUID Method

## Overview
Generates a cryptographically secure, version 4 Universally Unique Identifier (UUID).

## Syntax
```asp
result = crypto.UUID()
```

## Return Values
Returns a **String** containing a randomly generated UUID in the standard 8-4-4-4-12 format (e.g., `550e8400-e29b-41d4-a716-446655440000`).

## Remarks
The UUID method uses a cryptographically secure random number generator (CSPRNG) to ensure high entropy and prevent collisions. It is ideal for unique identifiers, session keys, and database primary keys.

## Code Example
```asp
<%
Dim crypto, newID
Set crypto = Server.CreateObject("G3CRYPTO")
newID = crypto.UUID()
Response.Write "New UUID: " & newID
Set crypto = Nothing
%>
```
