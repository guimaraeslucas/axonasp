
<%class cls_fileexplorer
public rootfolder
public thumbnailview
public expandall
private e
private fso
private folderpath
private doSF
private totalfoldersize
private allowedFileTypesforThumbingCopy
public bResize
Private Sub Class_Initialize
set fso=server.CreateObject ("scripting.filesystemobject")
set allowedFileTypesforThumbingCopy=allowedFileTypesforThumbing
bResize=false
end sub
Private Sub Class_Terminate
set fso=nothing
set allowedFileTypesforThumbingCopy=nothing
end sub
public function show()
dim rootfolderObj
set rootfolderObj=fso.GetFolder(rootfolder)
start(rootfolderObj)
if not bResize then
'show=replace(e,"<div></div>","",1,-1,1)
'show="<div style='position:relative;left:7px;margin: 2px 0px 6px 0px'><font style='font-size:smaller'>"&l("totalfoldersize")&": "& round(rootfolderObj.size/1024/1024,2) &" MB</font></div>" & show
end if
set rootfolderObj=nothing
end function
private function start(root)
showFolder(root)
end function
private function showFolder(folder)
folderpath=replace(folder.path,rootfolder,"",1,-1,1) & "/"
folderpath=replace(folderpath,"\","/",1,-1,1)
folderpath=C_VIRT_DIR & Application("QS_CMS_userfiles") & folderpath
folderpath=replace(folderpath,"//","/",1,-1,1)
'doSF=false
'e=e & "<div style='float:left;left:26px;clear:both;position:relative;width:100%'>" 
'e=e & "<div style='left:-13px;position:relative'>"
'e=e & "<a href=" & """" & "bs_fileexplorer.asp"
'if Request.QueryString ("doSF")<>folderpath and not expandall then
'e=e & "?doSF=" & quotRep(folderpath)&""""&"><img src='"&C_DIRECTORY_QUICKERSITE&"/fixedImages/plus2.gif' alt='' class='iM' />&nbsp;" 
'else
'e=e & """"&"><img class='iM' src='"&C_DIRECTORY_QUICKERSITE&"/fixedImages/minus2.gif' alt='' />&nbsp;" 
doSF=true
'end if
'e=e& "<b>" & folder.name & "</b></a>"
'e=e & "&nbsp;<a title=" & """" & l("newfolder") & """" & " href="& """" & "bs_fileexplorerAddFolder.asp?folderpath="&quotRep(folderpath)&"""" & "><img class='Im' alt='"&l("newfolder")&"' src='"&C_DIRECTORY_QUICKERSITE&"/fixedImages/twotone/news.gif' /></a>"
'dim aspuploader!
'e=e & "&nbsp;<a title=" & """" & l("newfile") & """" & " href="& """" & "bs_fileexplorerAddFile.asp?folderpath="&quotRep(folderpath)&"""" & "><img  class='Im' alt='"&l("newfile")&"' src='"&C_DIRECTORY_QUICKERSITE&"/fixedImages/twotone/upload.gif' /></a>"
'e=e & "&nbsp;<a title=" & """" & l("newfile") & """" & " href="& """" & "bs_fileexplorerAddFileNew.asp?folderpath="&quotRep(folderpath)&"""" & "><img  class='Im' alt='"&l("newfile")&"' src='"&C_DIRECTORY_QUICKERSITE&"/fixedImages/twotone/upload.gif' /></a>"
'if "/" & lcase(folder.name) & "/" <> lcase(Application("QS_CMS_userfiles")) then 
'e=e & "&nbsp;<a title="& """" & l("delete") & """" & " href='#' onclick=" & """" & "javascript: return delFol('" & quotRepJS(folderpath) & "');" & """" & "><img class='Im' alt='"&l("delete")&"' src='"&C_DIRECTORY_QUICKERSITE&"/fixedImages/twotone/trash.gif' /></a>"
'end if
'e=e & "</div>"
if doSF then showFiles folder,folderpath
dim subfolders, sf
set subfolders=folder.subfolders
for each sf in subfolders
start(sf)
next 
set subfolders=nothing
'e=e & "</div>"
'if not bResize then e=""
end function
private function showFiles(folder,path)
dim simplepath
simplepath=replace(path,C_VIRT_DIR & Application("QS_CMS_userfiles"),"",1,-1,1) 'needed for showthumb.aspx
dim files, file, filename, filesize
set files=folder.files
for each file in files
filesize=file.size
filename=file.name
next
set files=nothing
end function
public function deleteFile(filepath)
deleteFile=false
if not cleanUpPath(filepath) then endProcess
if fso.FileExists (server.mappath(filepath)) then
on error resume next
fso.DeleteFile(server.mappath(filepath))
if err.number=0 then deleteFile=true
on error goto 0
end if
end function
public function deletefolder(folderpath)
deletefolder=false
if not cleanUpPath(folderpath) then endProcess
if fso.FolderExists (server.mappath(folderpath)) then
on error resume next
fso.DeleteFolder(server.mappath(folderpath))
if err.number=0 then deletefolder=true
on error goto 0
end if
end function
public function addFolder(parentFolder,newFolder)
addFolder=false
if not cleanUpPath(parentFolder) then endProcess
if isValidFolderName(newFolder) then
if not fso.FolderExists (server.mappath(parentFolder&newFolder)) then
on error resume next
fso.CreateFolder (server.mappath(parentFolder&newFolder))
if err.number=0 then addFolder=true
on error goto 0
end if
end if
end function
public function cleanUpPath(path)
cleanUpPath=true
if lcase(left(path,len(C_VIRT_DIR&Application("QS_CMS_userfiles"))))<>lcase(C_VIRT_DIR&Application("QS_CMS_userfiles")) then
cleanUpPath=false
end if
end function
public sub endProcess
Response.Write "not allowed"
Response.End 
end sub
public function tempfolder
dim foldername
foldername=generatePassword&generatePassword&generatePassword
set tempfolder=fso.CreateFolder(rootfolder & "\" & foldername)
end function
public function SelectBoxFolders(root,selected)
SelectBoxFolders=getSelectBoxFolders(root)
'post processing
SelectBoxFolders=replace(SelectBoxFolders,root,"",1,-1,1)
SelectBoxFolders=replace(SelectBoxFolders,"\","/",1,-1,1)
'set selected
SelectBoxFolders=replace(SelectBoxFolders,"<option></option>","<option value=''>" & l("pleaseselect") & ":</option>",1,-1,1)
SelectBoxFolders=replace(SelectBoxFolders,"<option>" & selected & "</option>","<option selected=""selected"">" & selected & "</option>",1,-1,1)
end function
private function getSelectBoxFolders(root)
dim folder, subfolder
set folder=fso.GetFolder(root) 
getSelectBoxFolders=getSelectBoxFolders & "<option>" & folder.path & "</option>"
if folder.subfolders.count>0 then
for each subfolder in folder.subfolders
getSelectBoxFolders=getSelectBoxFolders&getSelectBoxFolders(subfolder.path)
next
end if
set folder=nothing
end function
end class%>
