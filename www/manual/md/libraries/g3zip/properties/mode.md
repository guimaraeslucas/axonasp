# Mode Property

## Overview
Indicates the current operating mode of the G3Pix AxonASP G3ZIP library instance.

## Syntax
```asp
currentMode = zip.Mode
```

## Return Values
Returns a **String**:
- **"r"**: The object is in **Read** mode (opened an existing archive).
- **"w"**: The object is in **Write** mode (created a new archive).
- **""** (Empty String): The object is currently closed.

## Remarks
- This property is read-only.
- The mode is automatically set by the **Open** and **Create** methods.

## Code Example
```asp
<%
Dim zip
Set zip = Server.CreateObject("G3ZIP")
zip.Create "/temp/test.zip"
If zip.Mode = "w" Then
    Response.Write "Ready to add files."
End If
Set zip = Nothing
%>
```
