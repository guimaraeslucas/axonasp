# Hash Property

## Overview
Returns the raw byte array of the most recently computed cryptographic hash.

## Syntax
```asp
byteArray = crypto.Hash
```

## Return Values
Returns an **Array** (VBScript Byte Array) containing the raw binary data of the last hash operation.

## Remarks
This property is useful when you need to store or transmit the hash in binary format rather than a hexadecimal string. It is updated automatically every time a hashing method (like **SHA256**, **MD5**, or **ComputeHash**) is called.

## Code Example
```asp
<%
Dim crypto, rawHash, i
Set crypto = Server.CreateObject("G3CRYPTO")
crypto.SHA256("Binary Output Test")
rawHash = crypto.Hash

Response.Write "First 4 bytes: "
For i = 0 To 3
    Response.Write Hex(rawHash(i)) & " "
Next
Set crypto = Nothing
%>
```
