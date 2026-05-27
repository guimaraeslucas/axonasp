# G3FILEUPLOADER Methods

## Overview
This page summarizes methods exposed by **G3FILEUPLOADER** in G3Pix AxonASP for upload validation, processing, and metadata inspection.

## Methods Reference

| Method | Returns | Description |
|---|---|---|
| AllowExtension | Empty | Adds one extension to the allowed list. |
| AllowExtensions | Empty | Adds multiple extensions to the allowed list from a comma-separated string. |
| BlockExtension | Empty | Adds one extension to the blocked list. |
| BlockExtensions | Empty | Adds multiple extensions to the blocked list from a comma-separated string. |
| Form | String | Gets a multipart form field value by name. |
| FormValue | String | Alias of Form. |
| IsValidExtension | Boolean | Validates whether an extension is currently allowed under configured rules. |
| GetFileInfo | Dictionary | Returns metadata for one uploaded form field. |
| GetAllFilesInfo | Array of Dictionary | Returns metadata for all uploaded files in the current request. |
| Process | Dictionary | Processes one uploaded file and saves it to disk. |
| Save | Dictionary | Alias of Process. |
| ProcessAll | Array of Dictionary | Processes all uploaded files and saves them to disk. |
| SaveAll | Array of Dictionary | Alias of ProcessAll. |
| SetUseAllowedOnly | Empty | Enables or disables allow-list-only validation mode. |

## Remarks
- Method names are case-insensitive.
- **Status Dictionary:** Processing methods always return a Dictionary containing `IsSuccess` (Boolean) and `ErrorMessage` (String) keys, along with metadata such as `FinalPath` and `RelativePath` on success.
