# Read and Write AxonASP Library Properties

## Overview
This page documents property surfaces extracted from AxonASP lib_*.go dispatch implementations.

## Syntax
```asp
Set obj = Server.CreateObject("ProgID")
value = obj.PropertyName
obj.PropertyName = newValue
`````

## Parameters and Arguments
- G3Crypto properties: hash, hashsize, bcryptcost, canreusetransform.
- G3DB properties: isopen, driver, dsn, lasterror, eof, bof, fields, count, committed, closed, lastinsertid, rowsaffected.
- G3FileUploader properties: blockedextensions, allowedextensions, maxfilesize, preserveoriginalname, debugmode.
- G3Image properties: hascontext, width, height, lasterror, lastmimetype, lasttempfile, lastbytes, defaultformat, jpgquality, alignleft, aligncenter, alignright, fillrulewinding, fillruleevenodd, linecapround, linecapbutt, linecapsquare, linejoinround, linejoinbevel.
- G3Mail properties: host, port, username, password, from, fromname, to, cc, bcc, subject, body, htmlbody, ishtml, bodyformat.
- G3PDF properties: lasterror, page, x, y, w, h, version.
- G3TAR properties: path, mode, count, lasterror.
- G3Zip properties: path, mode, count.
- G3ZLIB properties: lasterror.
- G3ZSTD properties: lasterror, level.
- WScript process properties: status, exitcode, processid, stdout, stderr, atendofstream, line.
- ADOX properties: tables, activeconnection, name, type.
- MSWC properties: border, clickable, targetframe.
- Scripting.Dictionary properties: count, comparemode, item, key, keys, items.
- Scripting.FileSystemObject family properties: drives, file.name, file.path, file.size, file.type, file.attributes, file.datecreated, file.datelastaccessed, file.datelastmodified, file.shortname, file.shortpath, file.parentfolder, file.drive, file.isrootfolder, folder.name, folder.path, folder.files, folder.subfolders, folder.size, folder.type, folder.attributes, folder.datecreated, folder.datelastaccessed, folder.datelastmodified, folder.shortname, folder.shortpath, folder.parentfolder, folder.drive, folder.isrootfolder, textstream.atendofstream, textstream.line, textstream.column, drive.driveletter, drive.drivetype, drive.filesystem, drive.availablespace, drive.freespace, drive.totalsize, drive.volumename, drive.path, drive.rootfolder, drive.serialnumber, drive.sharename, drive.isready, collection.count.
- VBScript.RegExp properties: pattern, ignorecase, global, multiline, matches.count, matches.length, match.value, match.firstindex, match.length, match.submatches, submatch.value, submatch.length.

## Return Values
Property reads return the current object state. Property writes apply changes when writable and validated by dispatch code.

## Remarks
- Some properties are read-only.
- Some properties represent object handles instead of scalar values.
- Invalid assignments raise runtime errors from the VM.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3Image")
img.DefaultFormat = "png"
img.JPGQuality = 92
Response.Write img.DefaultFormat
%>
`````
