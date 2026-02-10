
<%Class cls_zipper
private zipper,fso
public supportZipper
public zipname,ziponly
public bTreatAsArt4
public treatASJSTemplate
Private Sub Class_Initialize
bTreatAsArt4=false
treatASJSTemplate=false
on error resume next
Err.Clear
set zipper = server.createobject("aspZip.EasyZIP")
if err.number <>0 then
supportZipper=false
else
supportZipper=true
end if
on error goto 0
End Sub
Private Sub Class_Terminate
on error resume next
Set zipper = nothing
on error goto 0
End Sub
Public Function unpack(zipfile,folder)
On Error Resume Next
unpack=false
if not supportZipper then
'connect to remote unzipper
dim oXMLHTTP,arr,arrF,fileArr,fa,sfileName
Set oXMLHTTP = server.createobject("MSXML2.ServerXMLHTTP")
oXMLHTTP.open "GET", "https://www.quickersite.com/unzip.asp?ziponly=" & ziponly & "&zipfile=" & server.urlencode(customer.sUrl & replace(zipname,C_VIRT_DIR,"",1,-1,1))
oXMLHTTP.send
arr=split(oXMLHTTP.responseText,vbcrlf)
dim considerJSTemplate
considerJSTemplate=false
if instr(convertStr(oXMLHTTP.responseText),"superfish")<>0 then
considerJSTemplate=true
end if
if ubound(arr)<>0 then
set fso=server.createObject("scripting.filesystemobject")
ziponly=replace(ziponly,".zip","",1,-1,1)
on error resume next
fso.createfolder(server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates"))
fso.createfolder server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates") & "\" & ziponly
fso.createfolder server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates") & "\" & ziponly & "\images"
if considerJSTemplate then
fso.createfolder server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates") & "\" & ziponly & "\css"
fso.createfolder server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates") & "\" & ziponly & "\css\images"
fso.createfolder server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates") & "\" & ziponly & "\js"
fso.createfolder server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates") & "\" & ziponly & "\js\superfish"
fso.createfolder server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates") & "\" & ziponly & "\js\superfish\js"
fso.createfolder server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates") & "\" & ziponly & "\js\superfish\images"
fso.createfolder server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates") & "\" & ziponly & "\js\superfish\css"
end if
on error goto 0
on error resume next
for arrF=lbound(arr) to ubound(arr)
if not isLeeg(arr(arrF)) then
oXMLHTTP.open "GET", "https://www.quickersite.com/" & arr(arrF)
oXMLHTTP.send
fileArr=split(arr(arrF),"/")
for fa=lbound(fileArr) to ubound(fileArr)
sfileName=fileArr(fa)
next
if instr(arr(arrF),"/js/superfish/css/")<>0 then
sfileName="js\superfish\css\" & sfileName
elseif instr(arr(arrF),"/js/superfish/js/")<>0 then
sfileName="js\superfish\js\" & sfileName
elseif instr(arr(arrF),"/js/superfish/images/")<>0 then
sfileName="js\superfish\images\" & sfileName
elseif instr(arr(arrF),"/js/superfish/")<>0 then
sfileName="js\superfish\" & sfileName
elseif instr(arr(arrF),"/js/")<>0 then
sfileName="js\" & sfileName
elseif instr(arr(arrF),"/css/images/")<>0 then
sfileName="css\images\" & sfileName
elseif instr(arr(arrF),"/css/")<>0 then
sfileName="css\" & sfileName
elseif instr(arr(arrF),"/images/")<>0 then
sfileName="images\" & sfileName
end if
if oXMLHTTP.status=200 then
SaveBinaryData server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates/") & "\" & ziponly & "\" & sfileName, oXMLHTTP.responseBody
end if
end if
next
on error goto 0
set oXMLHTTP=nothing
set fso=nothing
if convertBool(Request.form("bRemove")) then
'if fe.fileexists(C_VIRT_DIR & Application("QS_CMS_userfiles") & Request.Form ("zip")) then
fe.deleteFile C_VIRT_DIR & Application("QS_CMS_userfiles") & Request.Form ("zip")
'end if
end if
response.redirect("bs_createTemplate.asp?path=" & server.urlencode(server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates") & "/" & ziponly))
end if
else
if lcase(GetFileExtension(zipfile))="zip" then
dim checkZIP
checkZIP=replace(zipFile,".zip","",1,-1,1)
dim files,file,bCheck1,bCheck2
bCheck1=false
bCheck2=true
zipper.Debug = false
zipper.ZipFileName = zipfile
for files = 0 to zipper.GetZipItemsCount - 1
select case lcase(GetFileExtension(zipper.GetZipItem_FIleName(files)))
case "png","jpg","js","gif","html","htm","artx","css","txt","swf","xml","old","ico","jpeg"
'do nothing
case else
response.write zipper.GetZipItem_FIleName(files)
response.end
bCheck2=false
exit for
end select
if left(lcase(zipper.GetZipItem_FIleName(files)),7)="images\" then bTreatAsArt4=true
if instr(zipper.GetZipItem_FIleName(files),"superfish") then treatASJSTemplate=true
Next
if treatASJSTemplate then bTreatAsArt4=false
bCheck1=true
if bCheck1 and bCheck2 then
'check if folder exists
set fso=server.createobject("scripting.filesystemobject")
if bTreatAsArt4 then
if not fso.folderexists(folder) then
fso.createFolder(folder) 
end if
else
folder=server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates")
end if
dim jsfileArr,jsfa,zfilename
if treatASJSTemplate then
if not fso.folderexists(folder) then
fso.createFolder(folder) 
end if
jsfileArr=split(zipfile,"\")
for jsfa=lbound(jsfileArr) to ubound(jsfileArr)
zfilename=jsfileArr(jsfa)
next
zfilename=replace(zfilename,".zip","",1,-1,1)
if fso.folderexists(folder & "\" & zfilename) then
fso.deletefolder folder  & "\" & zfilename
end if
fso.createFolder folder  & "\" & zfilename
folder=folder & "\" & zfilename
end if
zipper.ZipFileName=zipfile
zipper.ArgsAdd("*.*")
zipper.ExtrOptions = 3
zipper.ExtrbaseDir = folder
zipper.UnZip
if err.number=0 then
if treatASJSTemplate then
if fso.fileExists(folder & "\easyZip.txt") then
fso.deleteFile(folder & "\easyZip.txt")
end if
if fso.folderExists(folder & "\site") then
dim siteFolder,zfolder,zfile
set siteFolder=fso.getFolder(folder & "\site")
for each zfile in siteFolder.files
fso.movefile zfile.path,folder & "\"
next
for each zfolder in siteFolder.subfolders
fso.movefolder zfolder.path,folder & "\"
next
siteFolder.delete()
set siteFolder=nothing
err.clear()
end if
end if
end if
if err.number=0 then
unpack=true
end if
end if
end if
end if
if not unpack then
set fso=server.CreateObject ("scripting.filesystemobject")
if fso.FileExists (zipfile) then
fso.DeleteFile (zipfile)
end if
set fso=nothing
end if
On Error Goto 0
End Function
End class%>
