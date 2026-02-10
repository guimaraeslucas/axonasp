<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bTemplates%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Setup)%><%=getBOSetupMenu(btn_template)%><%dim template
set template=new cls_template
if convertGetal(decrypt(request("itemplateID")))<>0 then
checkCSRF()
select case Request.QueryString ("a")
case "copy"
template.copy()
case "delete"
template.remove()
case "default"
customer.defaultTemplate=template.iId
customer.save()
end select
end if
set template=nothing
dim templates
set templates=customer.templates
dim bTemplatesCanBeRemoved
bTemplatesCanBeRemoved=true
if templates.count<=1  then
bTemplatesCanBeRemoved=false
end if%><p align="center"><%=getArtLink("bs_templateEdit.asp",l("newtemplate"),"","","")%></p><%if templates.count>0 then%><table align="center" cellpadding="4" cellspacing="0"><%dim templateKey
for each templateKey in templates%><tr><td style="border-top:1px solid #DDD"><a href="bs_templateEdit.asp?itemplateID=<%=encrypt(templateKey)%>"><%=templates(templatekey).sName%></a><%if convertGetal(templateKey)=convertGetal(customer.defaultTemplate) then%>&nbsp;<i>(default)</i><%else%>&nbsp;<i>(<a href="bs_templateList.asp?a=default&amp;<%=QS_secCodeURL%>&amp;iTemplateID=<%=encrypt(templatekey)%>"><%=l("setaasdefault")%></a>)</i><%end if%></td><td style="border-top:1px solid #DDD"><%=getIcon(l("edit"),"edit","bs_templateEdit.asp?itemplateID="& encrypt(templateKey),"","edit"&templatekey)%></td><td style="border-top:1px solid #DDD"><%=getIcon(l("preview"),"search","#","javascript:window.open('" & C_DIRECTORY_QUICKERSITE & "/default.asp?iId="&encrypt(getHomePage.iId)&"&amp;previewTemplate="&encrypt(templateKey)&"','"&encrypt(templateKey)&"');","search"&templatekey)%></td><td style="border-top:1px solid #DDD"><%=getIcon(l("copyItem"),"copyItem","bs_templateList.asp?a=copy&" & QS_secCodeURL & "&amp;itemplateID="&encrypt(templateKey),"javascript:return confirm('"& l("areyousuretocopy") &"');","copy"&templateKey)%></td><td style="border-top:1px solid #DDD"><%if bTemplatesCanBeRemoved and customer.defaultTemplate<>templatekey then%><%=getIcon(l("deleteItem"),"delete","bs_templateList.asp?a=delete&" & QS_secCodeURL & "&amp;itemplateID="&encrypt(templateKey),"javascript:return confirm('"& l("areyousure") &"');","delete"&templateKey)%><%end if%> </td><td style="border-top:1px solid #DDD"><i>iID: <%=templateKey%></i></td></tr><%next%></table><%else%><p align="center"><%=l("notemplate")%></p><%end if
dim zipper
set zipper=new cls_zipper
if customer.supportZipper  then%><p align="center"><%=getArtLink("bs_uploadzip.asp","Upload zipped template","","","")%></p><%else%><p align=center>This installation cannot unzip templates. Install a free ZIP component by <b><a href="http://www.xstandard.com/en/documentation/xzip/" target="_blank">XStandard</a></b> on this IIS webserver.</p><%end if
if bBrowseOnlineTemplates then%><p align="center"><%=getArtLink("bs_templateSearch.asp","Search for templates online","","","")%><a href="bs_templateSearch.asp"></p><%end if
if not isleeg(sNewTemplatesPath) then
dim sInstallPath
sInstallPath=Request.QueryString ("sInstallPath")
if not isLeeg(sInstallPath) then
installTemplate sInstallPath
Response.Redirect ("bs_templatelist.asp?fbMessage=fb_templateadded")
end if%><a name="ts"></a><form name="selecttemplate" action="bs_templatelist.asp#ts" method="post"><input type="hidden" name="postback" value="<%=true%>" /><table align="center"><tr><td align="center"><select name="look" onchange="javascript:selecttemplate.submit();"><option value="">Browse our templates by category...</option><%=templateCatList(Request.Form ("look"))%></select></td></tr><%if Request.Form ("look")<>"" then%><tr><td><%=showTemplateBox(Request.Form ("look"))%></td></tr><%end if%></table></form><%
end if
set zipper=nothing
function templateCatList(selected)
dim fso,folder,f
set fso=server.CreateObject ("scripting.filesystemobject")

if fso.folderExists(sNewTemplatesPath) then

set folder=fso.GetFolder (sNewTemplatesPath)
for each f in folder.subfolders
templateCatList=templateCatList& "<option "
if convertStr(selected)=convertStr(f.path) then
templateCatList=templateCatList & " selected='selected' "
end if
templateCatList=templateCatList& " value=""" & server.HTMLEncode (f.path) & """>" & f.name  &  "</option>"
next
set folder=nothing

end if

set fso=nothing
end function
function showTemplateBox(path)
dim fso,folder,f,counter
set fso=server.CreateObject ("scripting.filesystemobject")
if fso.folderexists(path) then
set folder=fso.GetFolder (path)
counter=0
showTemplateBox="<table border=""0"" cellpadding=""10"" cellspacing=""0""><tr>"
for each f in folder.subfolders
showTemplateBox=showTemplateBox & "<td align=""center""><a onclick=""javascript:mw=window.open('" & sNewTemplatesURL & "/" & folder.name & "/" & f.name & "/page.html','pop','width=950,height=680;top=10;left=10,resizable=yes,scrollbars=yes');mw.focus();return false;"" href=""#""><img border=""0"" src='" & sNewTemplatesURL & "/" & folder.name & "/" & f.name & "/sc.png' alt=""Click to preview!"" width='180' style='width:180px' /></a><br />" & f.name & " - <a onclick=""javascript:return confirm('Are you sure?');"" href=""bs_templateList.asp?sInstallPath=" & server.URLEncode (path & "\" & f.name) & """>Install</a></td>"
counter=counter+1
if counter=3 then
showTemplateBox=showTemplateBox&"</tr><tr>"
counter=0
end if
next
showTemplateBox=showTemplateBox & "</tr></table>"
showTemplateBox=replace(showTemplateBox,"<tr></tr>","",1,-1,1)
set folder=nothing
else
showTemplateBox="<p>" & path & " does not exist</p>"
end if
set fso=nothing
end function
function installTemplate(pathtoinstall)
dim fso,folder
set fso=server.CreateObject ("scripting.filesystemobject")
dim urlParts,uP,sUp, textfile
urlParts=split(pathtoinstall,"\")
for uP=lbound(urlParts) to ubound(urlParts)
sUp=urlParts(uP)
next
'create sample templates
set template=new cls_template
template.bSetFooterVar=true
template.bSetRSSLink=true
template.bSetContactVar=true
template.bSetCustomHL=true
template.sName=sUp
textfile=fso.OpenTextFile(pathtoinstall & "\page.html",1).ReadAll
template.sValue=textfile
dim newSiteBP,NScounter
newSiteBP=server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") &"templates/"& sUp)
NScounter=1
while fso.FolderExists (newSiteBP)
newSiteBP=newSiteBP&NScounter
NScounter=NScounter+1
wend
'check if templatefolder exists
dim tempFolder
tempFolder=server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") &"templates")
if not fso.folderexists(tempFolder) then fso.createfolder(tempFolder)
fso.CopyFolder pathtoinstall,newSiteBP
if NScounter=1 then
folder=C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates/"& sUp &"/"
else
folder=C_VIRT_DIR & Application("QS_CMS_userfiles") & "templates/"& sUp & NScounter-1 & "/"
end if
template.treatAsArtisteer(folder)
template.initWAP
template.initPrint
template.initMobile
template.initEmail
template.initPrint
template.save()
set template=nothing
set fso=nothing
end function%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
