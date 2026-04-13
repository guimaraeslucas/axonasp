# HashSize Property

## Overview
Returns the size of the last computed cryptographic hash in bits.

## Syntax
```asp
sizeBits = crypto.HashSize
```

## Return Values
Returns an **Integer** representing the number of bits in the hash digest.

## Remarks
The value returned depends on the algorithm used for the last operation. For example, it returns 256 for **SHA256** and 512 for **SHA512**.

## Code Example
```asp
<%
Dim crypto
Set crypto = Server.CreateObject("G3CRYPTO")
crypto.SHA256("Bit Size Test")
Response.Write "Hash Size: " & crypto.HashSize & " bits"
Set crypto = Nothing
%>
```
