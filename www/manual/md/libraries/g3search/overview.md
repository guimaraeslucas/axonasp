# G3SEARCH Object Overview

## Overview
The **G3SEARCH** object is a high-performance native library for the **G3Pix AxonASP** environment. It provides comprehensive document indexing and full-text search capabilities using the **Bluge** search engine.

This object allows developers to recursively scan directories for specific file types, index their contents and filenames into a persistent local index, and execute complex match queries to retrieve relevant documents.

## Syntax
```asp
Dim search
Set search = Server.CreateObject("G3SEARCH")
```
```javascript
var search = Server.CreateObject("G3SEARCH");
```

## Remarks
The **G3SEARCH** object is optimized for local file search scenarios, such as documentation portals, knowledge bases, or internal site searches. It requires write access to the directory specified in the **IndexPath** property to store the search index files.

Before creating the object, enable the library in `config/axonasp.toml`:

`[g3search]`
`g3search_enabled = true`

When `g3search.g3search_enabled` is `false`, AxonASP raises an explicit runtime error and the object cannot be used.

The indexing process is recursive and can be filtered by file extension using the **Extension** property.

The **Search** method returns a two-dimensional VBScript array where each row contains two values: filename and relevance score.

## Code Example
The following example demonstrates how to instantiate the **G3SEARCH** object and check its default configuration.

```vbscript
<%
    Dim search
    Set search = Server.CreateObject("G3SEARCH")
    
    Response.Write "Default Extension: " & search.Extension
    
    Set search = Nothing
%>
```
