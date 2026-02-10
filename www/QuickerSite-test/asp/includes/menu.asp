
<%class cls_menu
public menu
private replaceMenu
private parents
public cssType
public bOnline
private start
public iSubMenuRoot
public sub class_initialize
start=true
iSubMenuRoot=null
end sub
public function getMenu(pageObj)
if isNull(pageObj) then
set pageObj=new cls_page
end if
menu=""
dim sCacheString
sCacheString=QS_CMS_cacheMenu & "-" & cssType & "-" & convertGetal(pageObj.iId) & "-" & convertGetal(logon.contact.iStatus)
if application(sCacheString)="" or isEmpty(sCacheString) then
application(sCacheString)=cacheMenu(pageObj,false)
end if
menu=application(sCacheString)
select case convertStr(cssType)
case "0"
getMenu="<ul>" & menu & "</ul>"
case "2"
getMenu="<div id='sitemap'><ul>" & menu & "</ul></div>"
case "artmenu"
getMenu="<ul class='artmenu'>" & menu & "</ul>"
case else
getMenu="<div id='menu'><ul id='QS_menulist'>" & menu & "</ul></div>"
end select
end function
public function cacheMenu(pageObj,bIntranet)
if isNull(pageObj) then
set pageObj=new cls_page
end if
dim rs, sql, isAM
sql="select iId from tblPage where iListPageID is null and "& sqlCustId &" and bLossePagina="&getSQLBoolean(false)&" and bDeleted="&getSQLBoolean(false)&" and "& sqlNull("iParentid",pageObj.iId) &" and bOnline="&getSQLBoolean(true)&" "
sql=sql& " and bIntranet="&getSQLBoolean(bIntranet)&" "
sql=sql&" order by iRang asc "
set rs=db.execute(sql)
if convertStr(cssType)="artmenu" and not bIntranet then
isAm=true
else
isAm=false
end if
dim page
while not rs.eof
set page=new cls_page
page.pick(rs(0))
cacheMenu=cacheMenu&"<li>"
select case convertStr(cssType)
case "2"
cacheMenu=cacheMenu& page.getParentLink
case else
page.bam=isAm
if not isNull(iSubMenuRoot) then
if convertGetal(iSubMenuRoot)=convertGetal(page.iParentID) then page.forceArtisteerRoot=true
end if
cacheMenu=cacheMenu& page.getLink(true)
end select
cacheMenu=cacheMenu&"<ul>"
cacheMenu=cacheMenu&cacheMenu(page,bIntranet)
cacheMenu=cacheMenu&"</ul></li>"
rs.movenext
set page=nothing
wend
set rs=nothing
cacheMenu=replace(cacheMenu,"<ul></ul>","",1,-1,1)
end function
public function getIntranetMenu(pageObj)
menu=""
if isNull(pageObj) then
set pageObj=new cls_page
end if
dim isAM
if convertStr(cssType)="artmenu" then
isAM=true
else
isAM=false
end if
if not convertBool(customer.intranetUse) then
getIntranetMenu=""
exit function
end if
dim intranetAddOn
if logon.authenticatedIntranet then
if logon.contact.iStatus=cs_read or logon.contact.iStatus=cs_readwrite then
dim sCacheString
sCacheString=QS_CMS_cacheIntranetMenu & "-" & cssType & "-" & convertGetal(pageObj.iId)& "-" & convertGetal(logon.contact.iStatus)
if application(sCacheString)="" or isEmpty(application(sCacheString)) then
application(sCacheString)=cacheMenu(pageObj,true)
end if
menu=application(sCacheString)
end if
if convertBool(customer.intranetUseMyProfile) then
intranetAddOn=intranetAddOn&"<li><a href='default.asp?pageAction=profile'>"
if not isAM then
intranetAddOn=intranetAddOn & customer.intranetMyProfile 
else
intranetAddOn=intranetAddOn & customer.intranetMyProfile
end if
intranetAddOn=intranetAddOn & "</a></li>"
if customer.bUseAvatars then
intranetAddOn=intranetAddOn&"<li><a class=""QSPPAVATAR"" href=""default.asp?pageAction=avataredit"">Avatar</a></li>"
end if
end if
if logon.contact.getTper.count>0 or logon.contact.getBper.count or logon.contact.getLper.count then
intranetAddOn=intranetAddOn&"<li><a href='default.asp?pageAction=editsite'>"
intranetAddOn=intranetAddOn & customer.sLabelEditSite 
intranetAddOn=intranetAddOn & "</a></li>"
end if
if not isLeeg(customer.intranetLogOff) then 
intranetAddOn=intranetAddOn&"<li><a href='default.asp?" & QS_secCodeURL & "&amp;pageAction="&cLogOff &"' onclick="&""""&"javascript:return(confirm('"& l("areyousure") &"'));"&""""&">"
intranetAddOn=intranetAddOn& customer.intranetLogOff
intranetAddOn=intranetAddOn&"</a></li>"
end if
else
dim gihpi
gihpi=encrypt(getIntranetHomePage.iId)
menu="<li><a href='default.asp?iId="& gihpi &"' id='QS_VMENU" & gihpi & "'>"
menu=menu& customer.intranetName
menu=menu&"</a></li>"
end if
select case convertStr(cssType)
case "0"
getIntranetMenu="<ul>" & menu 
getIntranetMenu=getIntranetMenu & intranetAddOn
getIntranetMenu=getIntranetMenu & "</ul>"
case "2"
getIntranetMenu="<div id='sitemapI'><ul>" & menu 
getIntranetMenu=getIntranetMenu & intranetAddOn
getIntranetMenu=getIntranetMenu & "</ul></div>"
case "artmenu"
getIntranetMenu="<ul class='artmenu'>" & menu 
getIntranetMenu=getIntranetMenu & intranetAddOn
getIntranetMenu=getIntranetMenu & "</ul>"
case else
getIntranetMenu="<div id='menuI'><ul id='QS_intranetmenulist'>" & menu
getIntranetMenu=getIntranetMenu & intranetAddOn
getIntranetMenu=getIntranetMenu&"</ul></div>"
end select
'response.clear
'response.write getIntranetMenu
'response.end
end function
public function getBOMenu(pageObj,intranet)
if isNull(pageObj) then
set pageObj=new cls_page
end if
dim rs,sql
sql="select iId from tblPage "
sql=sql& " where "& sqlCustId &" and bLossePagina=" & getSQLBoolean(false) & " and bDeleted=" & getSQLBoolean(false) & " and " & sqlNull("iParentid",pageObj.iId)
sql=sql& " and iListPageID is null "
if intranet then
sql=sql& " and bIntranet="&getSQLBoolean(true)&" "
else
sql=sql& " and (bIntranet="&getSQLBoolean(false)&" or bIntranet is null) "
end if
sql=sql& " order by iRang asc"
set rs=db.execute(sql)
dim page
while not rs.eof
set page=new cls_page
page.pick(rs(0))
menu=menu&"<li><table cellpadding=""1"" cellspacing=""0""><tr><td valign=""bottom"">" & page.getClickLink(true) & " </td>"
if page.bHomepage then
menu=menu & "<td>" & getIcon(l("homepage"),"homepage","#","javascript:return false;","h"&page.iId) & "</td>"
end if
if page.bContainerPage then
menu=menu & "<td>" & getIcon(l("containeritem"),"cont","#","javascript:return false;","c"&page.iId)  & "</td>"
end if
if not isLeeg(page.sOrderby) then
menu=menu & "<td>" & getIcon(l("list"),"listpage","#","javascript:return false;","c"&page.iId)  & "</td>"
end if
if not isLeeg(page.sExternalURL) then
menu=menu & "<td>" & getIcon(l("openlink"),"link","#","window.open('"&replace(page.sExternalURLPrefix&page.sExternalURL,"&","&amp;",1,-1,1)&"');","l"&page.iId)& "</td>"
end if
if secondAdmin.bPagesAdd then
menu=menu & "<td>" & getIcon(l("addnewitem"),"addnewitem","bs_setupPage.asp?bIntranet="&intranet&"&amp;iParentid="& EnCrypt(page.iId) ,"","an"&page.iId)& "</td>"
end if
if secondAdmin.bPagesMove and secondAdmin.bPageOrder then
menu=menu& "<td>" & getIcon(l("down"),"down","bs_default.asp?"&QS_secCodeURL&"&amp;btnaction=MoveDown&amp;iId="& EnCrypt(page.iId),"","d"&page.iId) & "</td>"
menu=menu& "<td>" & getIcon(l("up"),"up","bs_default.asp?"&QS_secCodeURL&"&amp;btnaction=MoveUp&amp;iId="& EnCrypt(page.iId),"","u"&page.iId) & "</td>"
end if
if secondAdmin.bPagesAdd then
menu=menu&"<td>" & getIcon(l("copyitem"),"copyitem","bs_default.asp?"&QS_secCodeURL&"&amp;btnaction=Copy&amp;iId=" & EnCrypt(page.iId),"return copyItem();","ci"&page.iId) & "</td>" 
end if
if secondAdmin.bPagesPW then
if isLeeg(page.sPw) then
menu=menu&"<td>" & getIcon(l("managepw"),"unlock","bs_applyPw.asp?iId=" & EnCrypt(page.iId),"","ul"&page.iId) & "</td>"
else
menu=menu&"<td>" & getIcon(l("managepw"),"lock","bs_applyPw.asp?iId=" & EnCrypt(page.iId),"","lk"&page.iId) & "</td>"
end if
end if
if secondAdmin.bPagesMove then
menu=menu&"<td>" & getIcon(l("moveitem"),"move"&tdir,"bs_selectPage.asp?"&QS_secCodeURL&"&amp;btnaction=Move&amp;iId="& EnCrypt(page.iId),"","mv"&page.iId) &"</td>"
end if
menu=menu & "</tr></table>"
menu=menu&"<ul>"
getBOMenu page,intranet
menu=menu&"</ul></li>"
rs.movenext
set page=nothing
wend
set rs=nothing
menu=replace(menu,"<ul></ul>","",1,-1,1)
getBOMenu="<ul class=menu>" & menu & "</ul>"
end function
public function showParents(pageId,withLink)
dim page
set page=new cls_page
page.pick(pageId)
if isLeeg(page.iId) then exit function
if start then
parents=page.sTitle & " &gt; " & parents
start=false
else
if withLink then
parents=page.getParentLink & " &gt; " & parents
else
parents=page.sTitle & " &gt; " & parents
end if
end if
dim rs, sql
sql="select iParentid from tblPage where "& sqlCustId &" and bDeleted="&getSQLBoolean(false)&" and bOnline="&getSQLBoolean(true)&" and iId="& convertGetal(page.iId) 
set rs=db.execute(sql)
if not rs.eof then
showParents rs(0),withLink
end if
showParents=left(parents,len(parents)-6)
if showParents=page.sTitle then showParents=""
end function
public function getReplaceMenu(pageObj,byref subPages, byref origPage)
if isNull(pageObj) then
set pageObj=new cls_page
end if
dim rs, sql
sql="select iId from tblPage where "
sql=sql& sqlCustId &" and iListPageID is null and bDeleted="&getSQLBoolean(false)&" and bLossePagina="&getSQLBoolean(false)&" and " & sqlNull("iParentid",pageObj.iId)
if origPage.bIntranet then
sql=sql &" and bIntranet="&getSQLBoolean(true)&" "
else
sql=sql &" and (bIntranet="&getSQLBoolean(false)&" or bIntranet is null) "
end if
if not origPage.canGoOffline then
sql=sql &" and bOnline="&getSQLBoolean(true)&" "
end if
sql=sql&" order by iRang asc"
set rs=db.execute(sql)
dim page
while not rs.eof
set page=new cls_page
page.pick(rs(0))
if origPage.iId<>page.iId and not subPages.exists(page.iID) then
replaceMenu=replaceMenu&"<li"
if not page.bOnline then
replaceMenu=replaceMenu&" style=""color:" & MYQS_offlineLinkColor & """><i>"&l("offline")&"</i>: "
else
replaceMenu=replaceMenu& ">"
end if
replaceMenu=replaceMenu&"<table cellpadding=""1"" cellspacing=""0""><tr>"
replaceMenu=replaceMenu& "<td valign=""bottom"">" & page.sTitle & " </td>"
if page.bHomepage then
replaceMenu=replaceMenu & "<td>" & getIcon(l("homepage"),"homepage","#","","home"&page.iId) & "</td>"
end if
replaceMenu=replaceMenu&"<td>"& getIcon(l("insertitem"),"insertItem"&tdir,"bs_selectPage.asp?"&QS_secCodeURL&"&amp;btnaction=Insert&amp;iId=" & EnCrypt(origPage.iId) & "&amp;insertInto="& EnCrypt(page.iId),"return insertItem();","insert"&page.iId) & "</td>" 
replaceMenu=replaceMenu&"</tr></table><ul>"
getReplaceMenu page,subPages,origPage
replaceMenu=replaceMenu&"</ul></li>"
end if
rs.movenext
set page=nothing
wend
set rs=nothing
replaceMenu=replace(replaceMenu,"<ul></ul>","",1,-1,1)
getReplaceMenu="<ul class=menu>" & replaceMenu & "</ul>"
end function
public function lossePaginas(intranet)
set lossePaginas=server.CreateObject ("scripting.dictionary")
dim rs, page, sql
sql="select iId from tblPage "
sql=sql & " where "& sqlCustId &" and bLossePagina="&getSQLBoolean(true)&" and bDeleted="&getSQLBoolean(false)&" "
if intranet then
sql=sql& " and bIntranet="&getSQLBoolean(true)&" "
else
sql=sql& " and (bIntranet="&getSQLBoolean(false)&" or bIntranet is null) "
end if
sql=sql & " order by sTitle asc"
set rs=db.execute(sql)
while not rs.eof
set page=new cls_page
page.pick(rs(0))
lossePaginas.Add page.iID, page
rs.movenext
wend
set rs=nothing
end function
public function getPRMenu(pageObj,intranet,getTPer,getBPer,getLper)
if isNull(pageObj) then
menu=""
set pageObj=new cls_page
end if
dim rs,sql
sql="select iId from tblPage "
sql=sql& " where "& sqlCustId &" and bLossePagina="&getSQLBoolean(false)&" and bDeleted="&getSQLBoolean(false)&" and " & sqlNull("iParentid",pageObj.iId)
sql=sql& " and iListPageID is null "
if intranet then
sql=sql& " and bIntranet="&getSQLBoolean(true)&" "
else
sql=sql& " and (bIntranet="&getSQLBoolean(false)&" or bIntranet is null) "
end if
sql=sql& " order by iRang asc"
set rs=db.execute(sql)
dim page
while not rs.eof
set page=new cls_page
page.pick(rs(0))
menu=menu&"<li><table cellpadding=""1"" cellspacing=""0""><tr><td valign=""bottom""><strong>" & page.sTitle & ":</strong></td>"
menu=menu & "<td><input type='checkbox' value='iPageIDTitle" & page.iId & "' name='bTitleID' "
if getTPer.exists(page.iId) then menu=menu & "checked"
menu=menu & " />" & l("edittitle") & "&nbsp;"
if not convertBool(page.bContainerPage) then
menu=menu & "<input type='checkbox' value='iPageIDBody" & page.iId & "' name='bBodyID' "
if getBPer.exists(page.iId) then menu=menu & "checked"
menu=menu & "/>" & l("editpage")
end if
if not isLeeg(page.sOrderby) then
menu=menu & "<input type='checkbox' value='iPageIDLP" & page.iId & "' name='bLPID' "
if getLPer.exists(page.iId) then menu=menu & "checked"
menu=menu & "/>" & l("managearticles")
end if 
menu=menu & "</td>"
menu=menu & "</tr></table>"
menu=menu&"<ul>"
getPRMenu page,intranet,getTPer,getBPer,getLper
menu=menu&"</ul></li>"
rs.movenext
set page=nothing
wend
set rs=nothing
menu=replace(menu,"<ul></ul>","",1,-1,1)
getPRMenu="<ul class=menu>" & menu & "</ul>"
end function
end class%>
