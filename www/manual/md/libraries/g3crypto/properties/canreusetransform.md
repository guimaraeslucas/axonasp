# CanReuseTransform Property

## Overview
Indicates whether the current cryptographic transform object can be reused for multiple operations.

## Syntax
```asp
isReusable = crypto.CanReuseTransform
```

## Return Values
Returns a **Boolean** value. This property always returns **True** in the G3CRYPTO implementation.

## Remarks
This property is provided for compatibility with standard cryptographic object patterns. In AxonASP, the G3CRYPTO library is designed to handle multiple sequential hashing operations without requiring manual reset or re-instantiation.

## Code Example
```asp
<%
Dim crypto
Set crypto = Server.CreateObject("G3CRYPTO")
If crypto.CanReuseTransform Then
    Response.Write "Transform is reusable"
End If
Set crypto = Nothing
%>
```
