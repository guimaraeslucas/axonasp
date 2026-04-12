# Call AxonASP Library Methods

## Overview
This page documents callable method surfaces extracted from AxonASP lib_*.go dispatch implementations.

## Syntax
```asp
Set obj = Server.CreateObject("ProgID")
result = obj.MethodName(arg1, arg2)
`````

## Parameters and Arguments
- G3Axon methods: axenginename, axversion, axgetenv, axgetconfig, axgetconfigkeys, axshutdownaxonaspserver, axchangedir, axcurrentdir, axhostnamevalue, axclearenvironment, axenvironmentlist, axenvironmentvalue, axprocessid, axeffectiveuserid, axdirseparator, axpathlistseparator, axintegersizebytes, axplatformbits, axexecutablepath, axexecute, axsysteminfo, s, n, v, m, axcurrentuser, axw, axmax, axmin, axintegermax, axintegermin, axceil, axfloor, axrand, axnumberformat, axpi, axsmallestfloatvalue, axfloatprecisiondigits, axcount, axexplode, aximplode, axarrayreverse, axrange, axstringreplace, axpad, axrepeat, axucfirst, axwordcount, axnl2br, axtrim, axstringgetcsv, axmd5, axsha1, axhash, sha256, sha1, md5, axbase64encode, axbase64decode, axurldecode, axrawurldecode, axrgbtohex, axhextorgb, axhtmlspecialchars, axstriptags, axfiltervalidateip, axfiltervalidateemail, axisint, axisfloat, axctypealpha, axctypealnum, axempty, axisset, axtime, axdate, axlastmodified, axgetremotefile, axgenerateguid, axgetdefaultcss, axgetlogo.
- G3Crypto methods: uuid, hashpassword, verifypassword, setbcryptcost, getbcryptcost, randombytes, randomhex, randombase64, md5, sha1, sha256, sha384, sha512, sha3_256, sha3_512, blake2b256, blake2b512, md5bytes, sha1bytes, sha256bytes, sha384bytes, sha512bytes, sha3_256bytes, sha3_512bytes, hmacsha256, hmacsha512, pbkdf2sha256, computehash, initialize.
- G3DB methods: open, openfromenv, close, query, queryrow, exec, prepare, begin, begintx, setmaxopenconns, setmaxidleconns, setconnmaxlifetime, setconnmaxidletime, stats, geterror, movenext, getrows, columns, scan, scanmap, commit, rollback, lastinsertid, rowsaffected.
- G3FC methods: create, extract, list, info, find, extractsingle.
- G3FILES methods: exists, read, write, append, delete, copy, move, size, mkdir, list, normalizeeol, converttextencoding, convertfileencoding.
- G3FileUploader methods: blockextension, allowextension, blockextensions, allowextensions, setuseallowedonly, process, processall, getfileinfo, getallfilesinfo, form, isvalidextension.
- G3HTTP methods: fetch.
- G3Image methods: close, new, loadimage, loadpng, loadjpg, newcontextforimage, savepng, savejpg, sethexcolor, setcolor, clear, setlinewidth, drawline, drawrectangle, drawcircle, drawellipse, stroke, fill, fillpreserve, strokepreserve, loadfontface, drawstring, drawstringanchored, measurestring, drawimage, renderviatemp.
- G3JSON methods: parse, stringify, newobject, newarray, loadfile.
- G3Mail methods: addaddress, addcc, addbcc, send, clear.
- G3PDF methods: new, addpage, close, output, I, F, setfont, setfontsize, settextcolor, setdrawcolor, setfillcolor, setlinewidth, setmargins, setleftmargin, settopmargin, setrightmargin, setx, sety, setxy, getx, gety, ln, cell, multicell, write, text, line, rect, image, addlink, setlink, link, settitle, setauthor, setsubject, setkeywords, setcreator, aliasnbpages, setdisplaymode, setcompression, writehtml, writehtmlfile, getpagewidth, getpageheight, getstringwidth.
- G3TAR methods: create, open, addfile, addfolder, addfiles, addtext, list, extractall, extractfile, getinfo, close.
- G3Template methods: render.
- G3Zip methods: open, create, addfile, addfolder, addtext, extractall, extractfile, list, getinfo, close.
- G3ZLIB methods: compress, decompress, decompresstext, compressmany, decompressmany, compressfile, decompressfile, clear.
- G3ZSTD methods: setlevel, compress, decompress, decompresstext, compressmany, decompressmany, compressfile, decompressfile, clear.
- WScript.Shell family methods: run, exec, createobject, getenv, expandenvironmentstrings, waituntildone, terminate, read, readline, readall, close.
- ADOX.Catalog family methods: tables.item, tables.count.
- MSWC methods: getadvertisement, getlistcount, getlistindex, getnextdescription, getnexturl, getpreviousdescription, getpreviousurl, getnthdescription, getnthurl, choosecontent, getallcontent, get, increment, remove, set, fileexists, owner, pluginexists, processform, hits, pagehit, reset, hasaccess.
- Scripting.Dictionary methods: add, exists, remove, removeall, keys, items, item, key, count.
- Scripting.FileSystemObject family methods: buildpath, copyfile, copyfolder, createfolder, createtextfile, deletefile, deletefolder, driveexists, fileexists, folderexists, getabsolutepathname, getbasename, getdrive, getdrivename, getextensionname, getfile, getfilename, getfileversion, getfolder, getparentfoldername, getspecialfolder, getstandardstream, gettempname, movefile, movefolder, opentextfile, file.copy, file.delete, file.move, file.openastextstream, folder.copy, folder.createtextfile, folder.delete, folder.move, textstream.read, textstream.readline, textstream.readall, textstream.write, textstream.writeline, textstream.writeblanklines, textstream.skip, textstream.skipline, textstream.close, collection.item.
- VBScript.RegExp family methods: execute, test, replace, matches.item, matches.count, match.submatches, submatches.item, submatches.count, submatch.value.

## Return Values
Method calls return operation-specific values such as Boolean, String, Integer, Array, Dictionary/native object handles, or Empty for compatibility setters and mutating operations.

## Remarks
- Unsupported method names raise object/member runtime errors.
- Some methods are object-family specific (for example ADODB Connection versus Recordset).
- See ADODB and MSXML pages for detailed compatibility method groups.

## Code Example
```asp
<%
Dim z
Set z = Server.CreateObject("G3Zip")
Call z.Create(Server.MapPath("./output.zip"))
Call z.AddText("notes.txt", "AxonASP")
Call z.Close()
%>
`````
