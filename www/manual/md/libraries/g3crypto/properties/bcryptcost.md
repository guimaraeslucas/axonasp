# BCryptCost Property

## Overview
Gets or sets the work factor (cost) used for bcrypt password hashing.

## Syntax
```asp
' Get current cost
cost = crypto.BCryptCost

' Set new cost
crypto.BCryptCost = 12
```

## Return Values
Returns an **Integer** (int64) representing the current work factor. The default value is 10.

## Remarks
The work factor represents the number of iterations performed during hashing (2^cost). Increasing this value improves security against brute-force attacks but significantly increases the CPU time required for each hash and verification. Accepted values are between 4 and 31.

## Code Example
```asp
<%
Dim crypto
Set crypto = Server.CreateObject("G3CRYPTO")
crypto.BCryptCost = 12
Response.Write "New BCrypt Cost: " & crypto.BCryptCost
Set crypto = Nothing
%>
```
