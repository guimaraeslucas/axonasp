
<%class cls_IIS
Public IISOBJ
Private siteRunner,siteObject,webobj
Private p_allsites
public logger
public overruleSI
Private Sub Class_Initialize
On Error Resume Next
Set IISOBJ = getObject("IIS://127.0.0.1/W3SVC")
set p_allsites=nothing
set logger=server.createobject("scripting.dictionary")
overruleSI=null
On Error Goto 0
end sub
Private sub Class_Terminate
On Error Resume Next
Set IISOBJ = nothing
set p_allsites=nothing
On Error Goto 0
end sub
Public Function GetNewSiteNumber()
on error resume next
if convertGetal(overruleSI)<>0 then
GetNewSiteNumber=overruleSI*100
else
dim rs
set rs=db.execute("select count(*) from tblCustomer")
dim counter
counter=convertGetal(rs(0))
counter=counter+1000
set rs=nothing
while GetAllSites.exists(cstr(counter))
counter=counter+1
wend
GetNewSiteNumber=counter
dumpError "GetNewSiteNumber",err
end if
on error goto 0
end function
Public Function GetAllSites
if p_allsites is nothing then
set p_allsites=server.CreateObject ("scripting.dictionary")
on error resume next
For each siteObject in IISOBJ
if (siteObject.Class = "IIsWebServer") then
set webobj = GetObject("IIS://localhost/w3svc/" & siteObject.name)
p_allsites.Add webobj.name,webobj
set webobj=nothing
end if
next
dumpError "GetAllSites",err
on error goto 0
end if
set GetAllSites=p_allsites
end function
end class
'response.write server.mappath("../")
'response.write server.mappath(C_DIRECTORY_QUICKERSITE & "/")
class cls_iisSite
Public sHostname,sPath,sSiteName,sitePath,QS_path,siteObject,basePath,SiteiId,otherbindings,samplesitepath,copyFromCustomerID,bDatabasehack
Private IISOBJ,siteRunner,CustomErrors,i,myIISOBJ,otherbindingsArr
Private Sub Class_Initialize
On Error Resume Next
Set myIISOBJ = new cls_IIS
set IISOBJ=myIISOBJ.IISOBJ 
dumpError "myIISOBJ.IISOBJ",err
QS_path="D:\inetpub\wwwroot\siteBuilder\"
'QS_path=server.mappath(C_DIRECTORY_QUICKERSITE & "/") & "\"
basePath="D:\inetpub\wwwroot\"
'basePath=server.mappath("../") & "\"
samplesitepath="emptysite.quickersite.com"
bDatabasehack=true
On Error Goto 0
end sub
Private Sub Class_Terminate
On Error Resume Next
Set myIISOBJ = nothing
set IISOBJ=nothing
On Error Goto 0
end sub
public Function IsValidSiteUrl(str)
Dim re
Set re = New RegExp
re.Pattern = "^(http|https)\://[a-zA-Z0-9\-\.]+\.[a-zA-Z]{2,3}(:[a-zA-Z0-9]*)?/?([a-zA-Z0-9\-\._\?\,\'/\\\+&amp;%\$#\=~])*[^\.\,\)\(\s]$"
IsValidSiteUrl = re.Test(str)
Set re = Nothing
End Function
public function updateServerBindings (mainBinding,allBindings)
on error resume next
if trim(mainBinding)="" or trim(allBindings)="" then
exit function
end if
dim GetAllSites
set GetAllSites=myIISOBJ.GetAllSites
for each siteObject in GetAllSites
dim fg
fg=GetAllSites(siteObject).ServerBindings
for i=lbound(fg) to ubound(fg)
if instr(lcase(fg(i)),lcase(replace(mainBinding,"http://","",1,-1,1)))<>0 then
dim sbArr,sbKey,sbDict
set sbDict=server.createObject("scripting.dictionary")
sbArr=split(allBindings,vbcrlf)
for sbKey=lbound(sbArr) to ubound(sbArr)
if not sbDict.exists(sbArr(sbKey)) then
'if IsValidSiteUrl(sbArr(sbKey)) then
if not isLeeg(sbArr(sbKey)) then
sbDict.add replace(lcase(trim(sbArr(sbKey))),"http://","",1,-1,1),""
end if
'end if
end if
next
 
if sbDict.count>0 then
dim cc
cc=sbDict.count-1
dim newSBS(),iSBS
redim newSBS(cc)
iSBS=0
for each sbKey in sbDict
newSBS(iSBS)=":80:" & sbKey
iSBS=iSBS+1
next
GetAllSites(siteObject).ServerBindings=newSBS
GetAllSites(siteObject).SetInfo()
end if
dumpError "IISOBJ.updateServerBindings "&newNumber,err
exit function
end if
next
next
set GetAllSites=nothing
dumpError "IISOBJ.updateServerBindings "&newNumber,err
on error goto 0
end function
Public function bHostnameExists
on error resume next
otherbindingsArr=split(otherbindings,vbcrlf)
bHostnameExists=false
if trim(sHostname)="" then
bHostnameExists=true
exit function
end if
sHostname=lcase(trim(sHostname))
'database hack
if bDatabaseHack then
dim rs, cHn,ssUrl,ssAlt
cHn=left(cleanUpStr(trim(sHostname)),200)
set rs=db.execute("select sUrl, sAlternateDomains from tblCustomer where iId<>" & convertGetal(SiteiId))
while not rs.eof
ssUrl=rs(0)
ssUrl=lcase(trim(convertStr(ssUrl)))
ssAlt=rs(1)
ssAlt=lcase(trim(convertStr(ssAlt)))
if not isLeeg(ssUrl) then
if instr(ssUrl,sHostname)<>0 then
bHostnameExists=true
exit function
end if
end if
if not isLeeg(ssAlt) then
if instr(ssAlt,sHostname)<>0 then
bHostnameExists=true
exit function
end if
end if
rs.movenext
wend
bHostnameExists=false
set rs=nothing
exit function
end if
'end database hack
dim GetAllSites
set GetAllSites=myIISOBJ.GetAllSites
for each siteObject in GetAllSites
dim fg
fg=GetAllSites(siteObject).ServerBindings
for i=lbound(fg) to ubound(fg)
if instr(lcase(fg(i)),sHostname)<>0 then
bHostnameExists=true
exit function
end if
next
next
set GetAllSites=nothing
on error goto 0
end function
Public function CreateFolder()
CreateFolder=false
dim fso
set fso=server.CreateObject ("scripting.filesystemobject")
if not fso.FolderExists (basePath&sPath) then
fso.CopyFolder basePath & samplesitepath,basePath&sPath
dim glo
if isleeg(copyFromCustomerID) then
glo="xxx"
else
glo=copyFromCustomerID
end if
'change global.asa
dim f
set f=fso.OpenTextFile (basePath&sPath&"\global.asa",1)
dim replaceXXX
replaceXXX=replace(f.readAll,glo,SiteiId,1,-1,1)
f.close
set f=nothing
dim newFile
set newFile=fso.CreateTextFile (basePath&sPath&"\global.asa",true)
newFile.writeline replaceXXX
newFile.close
set newFile=nothing
end if
set fso=nothing
CreateFolder=true
end function
Public function Create()
on error resume next
Create=true
if bHostnameExists then
create=false
exit function
end if
if not CreateFolder then
create=false
exit function
end if
if err.number <>0 then
create=false
exit function
end if
'create new site
dim newSite,newNumber
if convertGetal(SiteiId)<>0 then
myIISOBJ.overruleSI=SiteiId
end if
newNumber=myIISOBJ.GetNewSiteNumber
set newSite = IISOBJ.Create("IIsWebServer",newNumber)
dumpError "IISOBJ.Create "&newNumber,err
newSite.ServerComment	= left(sPath,40)
newSite.DefaultDoc	= "default.asp"
dim sbs(1),ij
if instr(sHostname,"www.")<>0 then
sbs(0)=":80:" & sHostname
sbs(1)=":80:" & replace(sHostname,"www.","",1,-1,1)
else
sbs(0)=":80:" & sHostname
sbs(1)=":80:" & "www." & sHostname
end if
newSite.ServerBindings=sbs
newSite.SetInfo()
'set custom 404 error message
CustomErrors = newSite.HttpErrors
For i = 0 To UBound(CustomErrors)
    If Left(CustomErrors(i),3) = "404" then
        CustomErrors(i) = "404,*,URL," & C_DIRECTORY_QUICKERSITE & "/default.asp"
        newSite.HttpErrors = CustomErrors
newSite.SetInfo()
        Exit For
     End If
Next
'set path
set sitePath = newSite.Create("IIsWebVirtualDir", "ROOT")
sitePath.KeyType="IIsWebVirtualDir"
sitePath.Path = basePath & sPath
sitePath.AccessRead = True
sitePath.AccessScript= True
sitePath.AppCreate2(2)
sitePath.AppFriendlyName = "Default Application"
sitePath.DontLog=true
sitePath.AuthFlags = 5 
sitePath.EnableDefaultDoc = True 
sitePath.DirBrowseFlags = &H4000003E 
sitePath.AccessFlags = 513 
sitePath.SetInfo()
newSite.Start()
if C_DIRECTORY_QUICKERSITE<>"" then
dim vDir
Set vDir = sitePath.Create("IIsWebVirtualDir", replace(C_DIRECTORY_QUICKERSITE,"/","",1,-1,1))
vDir.Path = QS_path
vDir.AccessRead = True
vDir.AccessScript= True
vDir.AccessWrite = False
vDir.EnableDirBrowsing = False
vDir.AppFriendlyName = "r"
vDir.SetInfo
set vDir=nothing
end if
set sitePath=nothing
set newSite=nothing
on Error goto 0
end function
public function remove
remove=false
if isLeeg(sHostname) then
exit function
end if
dim GetAllSites
set GetAllSites=myIISOBJ.GetAllSites
for each siteObject in GetAllSites
dim fg
fg=GetAllSites(siteObject).ServerBindings
for i=lbound(fg) to ubound(fg)
if instr(fg(i),sHostname)<>0 then
on error resume next
dim websiteParent
Set websiteParent = GetObject(GetAllSites(siteObject).Parent) 
websiteParent.Delete GetAllSites(siteObject).Class, GetAllSites(siteObject).Name 
dumpError "IISOBJ.DeleteSite",err
set websiteParent=nothing
if err.number=0 then 
remove=true
end if
on error goto 0
exit function
end if
next
next
set GetAllSites=nothing
end function
end class%>
