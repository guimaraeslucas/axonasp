# G3HTTP Methods

## Overview
This page provides a summary of the methods available in the **G3HTTP** library for performing outbound network requests in AxonASP.

## Method List

- **Fetch**: Sends an HTTP request to a remote server and returns the response body.
- **Request**: An alias for the **Fetch** method, providing the same functionality.

## Remarks
- Method names are case-insensitive.
- Both methods support optional HTTP verbs (such as POST, PUT, DELETE) and request body payloads.
- If the remote server responds with JSON, the methods automatically return a native object (Dictionary) or Array.
