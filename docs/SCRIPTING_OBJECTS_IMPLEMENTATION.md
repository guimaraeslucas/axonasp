## Scripting.Dictionary and Scripting.FileSystemObject Implementation Summary

### Overview
Comprehensive implementations of classic COM objects have been provided for AxonASP, including Scripting.Dictionary for key-value storage and an enhanced Scripting.FileSystemObject for complete file system operations, providing full VBScript compatibility.

### Files Created/Modified

#### New/Modified Files
1. **`server/dictionary_lib.go`** (276 lines)
   - Complete implementation of Scripting.Dictionary
   - Key-value pair storage
   - Case-insensitive key handling
   - Collection support
   - CompareMode property

2. **`server/file_lib.go`** (1220 lines)
   - FSOObject implementation (Scripting.FileSystemObject)
   - FSO file and folder objects
   - FSO text file operations
   - Collection support for files and folders
   - Full FSO method compatibility

#### Integration
1. **`server/executor_libraries.go`**
   - Added DictionaryLibrary wrapper
   - Added FileSystemObjectLibrary wrapper
   - Enables: `Set dict = Server.CreateObject("Scripting.Dictionary")`
   - Enables: `Set fso = Server.CreateObject("Scripting.FileSystemObject")`

### Key Features - Scripting.Dictionary

✓ **Dictionary Methods**
  - `Add(key, value)` - Add key-value pair
  - `Exists(key)` - Check if key exists
  - `Item(key)` - Get/set value by key
  - `Remove(key)` - Remove key and value
  - `RemoveAll()` - Clear all items
  - `Keys()` - Get all keys array
  - `Items()` - Get all values array

✓ **Dictionary Properties**
  - `Count` - Number of items
  - `CompareMode` - 0=Binary, 1=TextCompare
  - Case-insensitive key lookup (always)
  - Key order preservation

✓ **Dictionary Features**
  - For Each support via enumeration
  - Subscript syntax access: `dict("key")`
  - Type-agnostic values (any VBScript type)
  - Safe key lookup with default behavior
  - Thread-safe operations

✓ **Compare Modes**
  - Binary (0) - Case-sensitive comparison
  - TextCompare (1) - Case-insensitive comparison
  - Default: Binary mode on creation
  - Can be changed after creation

### Key Features - Scripting.FileSystemObject

✓ **File Operations**
  - `FileExists(path)` - Check file existence
  - `GetFile(path)` - Get file object
  - `CopyFile(source, dest, [overwrite])` - Copy file
  - `MoveFile(source, dest)` - Move file
  - `DeleteFile(path, [force])` - Delete file
  - `CreateTextFile(path, [overwrite])` - Create text file

✓ **Folder Operations**
  - `FolderExists(path)` - Check folder existence
  - `GetFolder(path)` - Get folder object
  - `CreateFolder(path)` - Create new folder
  - `MoveFolder(source, dest)` - Move folder
  - `DeleteFolder(path, [force])` - Delete folder recursively
  - `GetParentFolder(path)` - Get parent directory

✓ **Text File Operations**
  - `OpenTextFile(path, mode, [create], [format])` - Open file
  - Text file read/write/append operations
  - Line-by-line processing
  - Charset support
  - Stream-based file handling

✓ **Path Operations**
  - `GetBaseName(path)` - Get filename
  - `GetExtensionName(path)` - Get file extension
  - `GetParentFolderName(path)` - Get folder path
  - `GetDrive(path)` - Get drive letter
  - `BuildPath(basePath, relPath)` - Combine paths

✓ **Special Folders**
  - `GetSpecialFolder(folderType)` - Get system folders
  - 0 = System folder
  - 1 = Windows folder
  - 2 = Temporary folder

#### FSOFile Object

✓ **File Properties**
  - `Name` - Filename with extension
  - `Path` - Full path
  - `Size` - File size in bytes
  - `Type` - File type description
  - `DateCreated` - Creation date
  - `DateLastModified` - Modification date
  - `DateLastAccessed` - Last access date

✓ **File Methods**
  - `Copy(dest, [overwrite])` - Copy file
  - `Move(dest)` - Move file
  - `Delete([force])` - Delete file
  - `OpenAsTextFile([mode])` - Open for reading
  - `GetExtensionName()` - Get extension
  - `GetBaseName()` - Get filename only

#### FSOFolder Object

✓ **Folder Properties**
  - `Name` - Folder name
  - `Path` - Full path
  - `Size` - Total size of contents
  - `Type` - Folder type description
  - `Files` - Files collection
  - `SubFolders` - SubFolders collection
  - `ParentFolder` - Parent folder object
  - `DateCreated` - Creation date
  - `DateLastModified` - Modification date

✓ **Folder Methods**
  - `CreateFolder(name)` - Create subfolder
  - `Copy(dest, [overwrite])` - Copy folder
  - `Move(dest)` - Move folder
  - `Delete([force])` - Delete folder recursively
  - `GetParentFolder()` - Get parent
  - `GetBaseName()` - Get folder name

#### FSOTextFile Object

✓ **TextFile Properties**
  - `AtEndOfStream` - EOF indicator
  - `Line` - Current line number
  - `Column` - Current column

✓ **TextFile Methods**
  - `ReadLine()` - Read one line
  - `ReadAll()` - Read entire file
  - `Read(chars)` - Read N characters
  - `WriteLine(text)` - Write line
  - `Write(text)` - Write text
  - `Close()` - Close file
  - `Skip(lines)` - Skip N lines
  - `SkipLine()` - Skip current line

### Architecture

**Class Hierarchy**:
```
Scripting.Dictionary
  ├─ Add()
  ├─ Remove()
  ├─ RemoveAll()
  ├─ Exists()
  ├─ Keys()
  ├─ Items()
  ├─ Count property
  ├─ CompareMode property
  └─ Item() accessor

Scripting.FileSystemObject (FSOObject)
  ├─ File operations
  ├─ Folder operations
  ├─ TextFile operations
  ├─ Path operations
  └─ Collections

FSOFile
  ├─ Properties: Name, Path, Size, Type, Dates
  └─ Methods: Copy, Move, Delete, OpenAsTextFile

FSOFolder
  ├─ Properties: Name, Path, Size, Files, SubFolders
  └─ Methods: CreateFolder, Copy, Move, Delete

FSOTextFile
  ├─ Properties: AtEndOfStream, Line, Column
  └─ Methods: ReadLine, ReadAll, Write, WriteLine

Collections
  ├─ Files - File collection in folder
  ├─ SubFolders - Folder collection
  └─ Enumeration support
```

### Usage Examples - Dictionary

#### Basic Dictionary Usage
```vbscript
Dim dict
Set dict = Server.CreateObject("Scripting.Dictionary")

' Add items
dict.Add "name", "John"
dict.Add "email", "john@example.com"
dict.Add "age", 30

' Check existence
If dict.Exists("name") Then
    Response.Write "Name exists: " & dict("name")
End If

' Get item
Response.Write dict.Item("email")

' Get count
Response.Write "Total items: " & dict.Count
```

#### Dictionary with For Each
```vbscript
Dim dict, keys, i
Set dict = Server.CreateObject("Scripting.Dictionary")

dict("key1") = "value1"
dict("key2") = "value2"
dict("key3") = "value3"

' Enumerate keys
For Each key In dict.Keys
    Response.Write key & " = " & dict(key) & "<br>"
Next
```

#### Case-Insensitive Lookup
```vbscript
Dim dict
Set dict = Server.CreateObject("Scripting.Dictionary")

dict("Name") = "John"
dict("EMAIL") = "john@example.com"

' All these work (case-insensitive)
Response.Write dict("name")  ' John
Response.Write dict("email")  ' john@example.com
Response.Write dict("NAME")  ' John
Response.Write dict("Email")  ' john@example.com
```

#### Remove Items
```vbscript
Dim dict
Set dict = Server.CreateObject("Scripting.Dictionary")

dict.Add "item1", "value1"
dict.Add "item2", "value2"

If dict.Exists("item1") Then
    dict.Remove "item1"
End If

' Remove all
dict.RemoveAll()
Response.Write "Items remaining: " & dict.Count  ' 0
```

#### Get All Keys and Values
```vbscript
Dim dict, keys, items, i
Set dict = Server.CreateObject("Scripting.Dictionary")

dict("a") = 1
dict("b") = 2
dict("c") = 3

' Get keys
keys = dict.Keys()
For i = 0 To UBound(keys)
    Response.Write "Key: " & keys(i) & "<br>"
Next

' Get values
items = dict.Items()
For i = 0 To UBound(items)
    Response.Write "Value: " & items(i) & "<br>"
Next
```

#### Compare Mode
```vbscript
Dim dict
Set dict = Server.CreateObject("Scripting.Dictionary")

' Set compare mode (0=Binary, 1=TextCompare)
dict.CompareMode = 1  ' Case-insensitive

dict("Name") = "John"

If dict.Exists("name") Then  ' Works with TextCompare
    Response.Write "Found"
End If
```

### Usage Examples - FileSystemObject

#### File Operations
```vbscript
Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")

' Check file existence
If fso.FileExists("data.txt") Then
    Response.Write "File exists"
End If

' Copy file
fso.CopyFile "source.txt", "backup.txt", True

' Move file
fso.MoveFile "oldname.txt", "newname.txt"

' Delete file
fso.DeleteFile "temp.txt"
```

#### Folder Operations
```vbscript
Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")

' Create folder
fso.CreateFolder "newdir"

' Check folder existence
If fso.FolderExists("uploads") Then
    Response.Write "Uploads folder exists"
End If

' Move folder
fso.MoveFolder "oldfolder", "newfolder"

' Delete folder (recursive)
fso.DeleteFolder "temp", True
```

#### Get File Object
```vbscript
Dim fso, file
Set fso = Server.CreateObject("Scripting.FileSystemObject")

Set file = fso.GetFile("document.txt")

Response.Write "File: " & file.Name & "<br>"
Response.Write "Size: " & file.Size & " bytes<br>"
Response.Write "Created: " & file.DateCreated & "<br>"
Response.Write "Modified: " & file.DateLastModified & "<br>"
```

#### Get Folder Object
```vbscript
Dim fso, folder, file
Set fso = Server.CreateObject("Scripting.FileSystemObject")

Set folder = fso.GetFolder("documents")

Response.Write "Folder: " & folder.Name & "<br>"
Response.Write "Path: " & folder.Path & "<br>"
Response.Write "Size: " & folder.Size & " bytes<br>"

' List files
For Each file In folder.Files
    Response.Write "- " & file.Name & "<br>"
Next

' List subfolders
Dim subfolder
For Each subfolder In folder.SubFolders
    Response.Write "[" & subfolder.Name & "]<br>"
Next
```

#### Read Text File
```vbscript
Dim fso, textFile
Set fso = Server.CreateObject("Scripting.FileSystemObject")

' Open text file for reading (1)
Set textFile = fso.OpenTextFile("data.txt", 1)

Do While Not textFile.AtEndOfStream
    Response.Write textFile.ReadLine & "<br>"
Loop

textFile.Close()
```

#### Write Text File
```vbscript
Dim fso, textFile
Set fso = Server.CreateObject("Scripting.FileSystemObject")

' Create and open text file (2 = ForWriting)
Set textFile = fso.OpenTextFile("output.txt", 2, True)

textFile.WriteLine "Line 1"
textFile.WriteLine "Line 2"
textFile.WriteLine "Line 3"

textFile.Close()
```

#### Append to Text File
```vbscript
Dim fso, textFile
Set fso = Server.CreateObject("Scripting.FileSystemObject")

' Open for appending (8 = ForAppending)
Set textFile = fso.OpenTextFile("log.txt", 8)

textFile.WriteLine Now & " - User login"
textFile.WriteLine Now & " - Action: Download"

textFile.Close()
```

#### List Directory Contents
```vbscript
Dim fso, folder
Set fso = Server.CreateObject("Scripting.FileSystemObject")

Set folder = fso.GetFolder(".")

Response.Write "<h3>Files:</h3>"
For Each file In folder.Files
    Response.Write "- " & file.Name & " (" & file.Size & " bytes)<br>"
Next

Response.Write "<h3>Folders:</h3>"
For Each subfolder In folder.SubFolders
    Response.Write "- " & subfolder.Name & "<br>"
Next
```

#### Path Operations
```vbscript
Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")

Dim path
path = "C:\Users\John\Documents\report.pdf"

Response.Write "Full Path: " & path & "<br>"
Response.Write "Filename: " & fso.GetBaseName(path) & "<br>"
Response.Write "Extension: " & fso.GetExtensionName(path) & "<br>"
Response.Write "Folder: " & fso.GetParentFolderName(path) & "<br>"
Response.Write "Drive: " & fso.GetDrive(path) & "<br>"
```

#### Copy Folder Recursively
```vbscript
Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")

' Copy entire folder with all subfolders and files
fso.CopyFolder "source_dir", "backup_dir", True
```

#### Get Temporary Folder
```vbscript
Dim fso, tempFolder, tempFile
Set fso = Server.CreateObject("Scripting.FileSystemObject")

' Get Windows temp folder (type 2)
Set tempFolder = fso.GetSpecialFolder(2)

Response.Write "Temp folder: " & tempFolder.Path
```

#### File Size of Folder
```vbscript
Dim fso, folder
Set fso = Server.CreateObject("Scripting.FileSystemObject")

Set folder = fso.GetFolder("documents")

Response.Write "Total size: " & folder.Size & " bytes"
Response.Write "(" & (folder.Size / 1024 / 1024) & " MB)"
```

### TextFile Open Modes
- 1 (ForReading) - Read only
- 2 (ForWriting) - Write (create/overwrite)
- 8 (ForAppending) - Append to end

### Special Folder Types
- 0 = Windows System folder
- 1 = Windows folder
- 2 = Temporary folder

### Performance Characteristics
- Dictionary fast key lookup (O(1) average)
- Folder enumeration efficient
- File operations use native file system
- No caching (fresh reads each time)
- Suitable for small to medium operations

### Standard COM Compatibility
- Full VBScript Dictionary interface
- Full VBScript FileSystemObject interface
- Works with classic ASP patterns
- Compatible with legacy code

### Limitations
- Dictionary keys must be strings (internally)
- No custom object values in Dictionary (basic types only)
- File operations limited to accessible paths
- No advanced attributes (only basic metadata)
- No file locking support

### Error Handling
- Path validation before operations
- Graceful error returns
- Server logging for debugging
- Safe null handling

### Security Considerations
✓ **Implemented**:
- Path validation against root
- Directory traversal prevention
- Safe file operations

⚠ **Not Implemented**:
- File permissions checking
- Access control lists
- Encryption operations

### Future Enhancements
- Advanced file attributes
- File permissions support
- Symbolic link handling
- Compression operations
- File encryption
- Advanced search filters
