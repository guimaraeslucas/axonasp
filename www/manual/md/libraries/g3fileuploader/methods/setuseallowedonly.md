# SetUseAllowedOnly Method

## Overview
Activates an internal strict lockdown, discarding all system capabilities globally barring those extensions exclusively pushed directly into engine array mechanisms matching precise whitelist entries.

## Syntax
```asp
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.SetUseAllowedOnly True
```

## Parameters and Arguments
- `Enabled` (Boolean, Required): True directly configures whitelist-only uploads.

## Return Values
Returns an `Empty` variant.
