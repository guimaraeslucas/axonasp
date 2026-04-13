# Use the ASP Dictionary Object

## Overview
The **Dictionary** object in G3Pix AxonASP is a highly efficient data structure used to store and manage key-value pairs. It functions similarly to an associative array or a hash map, allowing for fast retrieval of data based on unique keys. This object is preferred over standard arrays for tasks involving lookup tables, JSON-like maps, and managing complex configuration data.

## Syntax
To instantiate the object, use the following syntax:
```asp
Set dict = Server.CreateObject("Scripting.Dictionary")
```

## How it Works
The AxonASP Virtual Machine provides a native implementation of the **Scripting.Dictionary** object. 
- **Keys and Items**: Each entry consists of a unique **Key** and an associated **Item** (value).
- **Type Flexibility**: Keys and items can be any Variant type, including strings, numbers, and other object references.
- **Comparison Mode**: The object can be configured to perform case-sensitive or case-insensitive key comparisons.

## API Reference

### Methods
- **Add(key, item)**: Adds a new key-value pair to the dictionary. Returns **Empty**.
- **Exists(key)**: Returns a **Boolean** indicating whether the specified key exists in the dictionary.
- **Items()**: Returns a **VBArray** containing all the items in the dictionary.
- **Keys()**: Returns a **VBArray** containing all the keys in the dictionary.
- **Remove(key)**: Removes a specific key-value pair from the dictionary. Returns **Empty**.
- **RemoveAll()**: Clears all entries from the dictionary. Returns **Empty**.

### Properties
- **CompareMode**: Gets or sets an **Integer** specifying the comparison mode (0 for Binary/Case-Sensitive, 1 for Text/Case-Insensitive).
- **Count**: Returns an **Integer** representing the number of entries currently in the dictionary.
- **Item(key)**: Gets or sets the value for a specified key. This is the default property.
- **Key(oldKey)**: Sets a new name for an existing key.

## Code Example
The following example demonstrates how to create a dictionary, add data, and iterate through the keys.

```asp
<%
Dim dict, k
Set dict = Server.CreateObject("Scripting.Dictionary")

' Set case-insensitive mode
dict.CompareMode = 1

' Populate data
dict.Add "Engine", "AxonASP"
dict.Add "Version", "2.0"
dict.Add "Status", "Running"

' Check for existence and retrieve item
If dict.Exists("Engine") Then
    Response.Write "Application Engine: " & dict("Engine") & "<br>"
End If

' Iterate through all entries
For Each k In dict.Keys
    Response.Write k & ": " & dict(k) & "<br>"
Next

Set dict = Nothing
%>
```

## Remarks
- **Unique Keys**: Attempting to add a key that already exists will trigger a runtime error. Use the **Exists** method to check before adding.
- **Resource Management**: Always set the dictionary object to **Nothing** once it is no longer needed to free system memory.
- **Performance**: The **Dictionary** object is optimized for fast lookups even with large datasets.
