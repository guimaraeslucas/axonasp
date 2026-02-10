
<%class cls_feed
Public iId, sName, sPrefixUrl, sUrl, iMaxItems, sKeywords, bOpenLinkInNW, iReloadSec, sCode, iLimitTo, iTitleLimitTo, iCache
Public bShowTitle, bLinkOnTitle, bShowAuthor, bShowCategory, bShowDate, sTemplate, bTemplate, sTemplateInit
Public bTemplateInit, sHTMLBefore, sHTMLAfter, bRandom, bEnableJS,overruleCID, sUrls
Private Sub Class_Initialize
On Error Resume Next
iId	= null
iMaxItems	= 10
bOpenLinkInNW	= true
iReloadSec	= 0
iLimitTo	= 0
iTitleLimitTo=0
sPrefixUrl	= "http://"
iCache	= 3600
bShowTitle	= true
bLinkOnTitle	= true
bShowAuthor	= true
bShowCategory	= true
bShowDate	= true
'sTemplateInit	= "<div class='QS_feeditem'>"&vbcrlf&"  <div class='QS_feeditemtitle'><a  target='_new'  href=""{LINK}"">{TITLE}</a></div>"&vbcrlf&"  <div class='QS_feeditemdetails'>{DATE} | {AUTHOR}</div>"&vbcrlf&"  <div class='QS_feeditemcategory'>{CATEGORY}</div>"&vbcrlf&"  <div class='QS_feeditemdescription'>{DESCRIPTION}</div>"&vbcrlf&"</div>"
bTemplate	= false
bRandom	= false
bEnableJS	= false
pick(decrypt(request("ifeedId")))
On Error Goto 0
end sub
Public property get sLink
sLink=sPrefixUrl&sUrl
end property
Public Function Pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblfeed where iCustomerID="&cid&" and iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
sPrefixUrl	= rs("sPrefixUrl")
sUrl	= rs("sUrl")
iMaxItems	= rs("iMaxItems")
sKeywords	= rs("sKeywords")
bOpenLinkInNW	= rs("bOpenLinkInNW")
iReloadSec	= rs("iReloadSec")
sName	= rs("sName")
sCode	= rs("sCode")
iLimitTo	= rs("iLimitTo")
iTitleLimitTo=rs("iTitleLimitTo")
iCache	= rs("iCache")
bShowTitle	= rs("bShowTitle")
bLinkOnTitle	= rs("bLinkOnTitle")
bShowAuthor	= rs("bShowAuthor")
bShowCategory	= rs("bShowCategory")
bShowDate	= rs("bShowDate")
sTemplate	= rs("sTemplate")
bTemplate	= rs("bTemplate")
sHTMLBefore	= rs("sHTMLBefore")
sHTMLAfter	= rs("sHTMLAfter")
bRandom	= rs("bRandom")
bEnableJS	= rs("bEnableJS")
sUrls	= rs("sUrls")
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
if isLeeg(sName) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(sUrl) and isLeeg(sUrls) then
check=false
message.AddError("err_mandatory")
else
sUrl=replace(sUrl,"http://","",1,-1,1)
sUrl=replace(sUrl,"https://","",1,-1,1)
if instr(surl,"default.asp?iId=")<>0 then
check=false
message.AddError("err_mandatory")
end if
if instr(surl,"default.asp?sCode=")<>0 then
check=false
message.AddError("err_mandatory")
end if
end if
if not isLeeg(sCode) then
dim checkRS
set checkRS=db.execute("select count(iId) from tblFeed where iId<>" & convertGetal(iId) & " and iCustomerID=" & cID & " and sCode='" & cleanup(ucase(sCode)) & "'")
if clng(checkRS(0))>0 then
check=false
message.AddError("err_doublefeed")
end if
set checkRS=nothing
end if
End Function
Public Function Save
if check() then
save=true
else
save=false
exit function
end if
dim rs
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblfeed where 1=2"
rs.AddNew
rs("dCreatedTS")=now()
else
rs.Open "select * from tblfeed where iId="& iId
end if
rs("sName")	= left(sName,50)
rs("sUrl")	= sUrl
rs("sPrefixUrl")	= sPrefixUrl
rs("iMaxItems")	= iMaxItems
rs("sKeywords")	= sKeywords
rs("bOpenLinkInNW")	= bOpenLinkInNW
rs("iReloadSec")	= iReloadSec
if isLeeg(overruleCID) then
rs("iCustomerID")	= cId
else
rs("iCustomerID")	= overruleCID
end if
rs("dUpdatedTS")	= now()
rs("sCode")	= sCode
rs("iLimitTo")	= iLimitTo
rs("iTitleLimitTo")=iTitleLimitTo
rs("iCache")	= iCache
rs("bShowTitle")	= bShowTitle
rs("bLinkOnTitle")	= bLinkOnTitle
rs("bShowAuthor")	= bShowAuthor
rs("bShowCategory")	= bShowCategory
rs("bShowDate")	= bShowDate
rs("sTemplate")	= sTemplate
rs("bTemplate")	= bTemplate
rs("sHTMLAfter")	= sHTMLAfter
rs("sHTMLBefore")	= sHTMLBefore
rs("bRandom")	= bRandom
rs("bEnableJS")	= bEnableJS
rs("sUrls")	= sUrls
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
customer.cacheFeeds()
application(QS_CMS_cacheFEED & iId)=""
end function
public function getRequestValues()
sName	= convertStr(Request.Form ("sName"))
sPrefixUrl	= convertStr(Request.Form ("sPrefixUrl"))
sUrl	= convertStr(Request.Form ("sUrl"))
iMaxItems	= convertGetal(Request.Form ("iMaxItems"))
sKeywords	= convertStr(Request.Form ("sKeywords"))
bOpenLinkInNW	= convertBool(Request.Form ("bOpenLinkInNW"))
iReloadSec	= convertGetal(Request.Form ("iReloadSec"))
sCode	= ucase(convertStr(Request.Form ("sCode")))
iLimitTo	= convertGetal(Request.Form ("iLimitTo"))
iTitleLimitTo=convertGetal(Request.Form("iTitleLimitTo"))
iCache	= convertGetal(Request.Form("iCache"))
bShowTitle	= convertBool(Request.Form ("bShowTitle"))
bLinkOnTitle	= convertBool(Request.Form ("bLinkOnTitle"))
bShowAuthor	= convertBool(Request.Form ("bShowAuthor"))
bShowCategory	= convertBool(Request.Form ("bShowCategory"))
bShowDate	= convertBool(Request.Form ("bShowDate"))
sTemplate	= convertStr(Request.Form ("sTemplate"))
bTemplate	= convertBool(Request.Form ("bTemplate"))
sHTMLBefore	= convertStr(Request.Form ("sHTMLBefore"))
sHTMLAfter	= convertStr(Request.Form ("sHTMLAfter"))
bRandom	= convertBool(Request.Form ("bRandom"))
bEnableJS	= convertBool(Request.Form ("bEnableJS"))
sUrls	= convertStr(Request.Form ("sUrls"))
if convertGetal(iId)=0 then
if bTemplate then
if convertBool(Request.Form ("postback")) then
if Request.Form ("btnaction")<>l("save") then
sTemplate=sTemplateInit
end if
end if
end if
end if
end function
public function remove
if not isLeeg(iId) then
dim rs
set rs=db.execute("update tblPage set ifeedID=null where ifeedID="& iId)
set rs=nothing
set rs=db.execute("delete from tblfeed where iId="& iId)
set rs=nothing
customer.cacheFeeds()
application(QS_CMS_cacheFEED & iId)=""
end if
end function
public function copy()
if isNumeriek(iId) then
iId=null
sName=l("copyof") & " " & sName
sCode=""
save()
end if
end function
public function build
build=""
if convertGetal(iId)<>0 then
if isEmpty(application(QS_CMS_cacheFEED & iId)) or application(QS_CMS_cacheFEED & iId)="" or cdate(now())>cdate(application(QS_CMS_cacheFEED & iId & "TS")) then
On Error Resume Next
dim feedType, ItemTemplate
if bTemplate then
ItemTemplate=treatConstants(sTemplate,true)
else
ItemTemplate="<div class='QS_feeditem'>"
if bShowTitle then
ItemTemplate=ItemTemplate&"<div class='QS_feeditemtitle'>"
if bLinkOnTitle then
ItemTemplate=ItemTemplate&"<a "
if bOpenLinkInNW then
ItemTemplate=ItemTemplate&" target='" & GeneratePassWord & "' "
end if
ItemTemplate=ItemTemplate& " href=" & """{LINK}""" & ">{TITLE}</a>"
else
ItemTemplate=ItemTemplate&"{TITLE}"
end if
ItemTemplate=ItemTemplate&"</div>"
end if
if bShowDate or bShowAuthor then
ItemTemplate=ItemTemplate&"<div class='QS_feeditemdetails'>"
if bShowDate then
ItemTemplate=ItemTemplate&"{DATE}&nbsp;|&nbsp;"
end if
if bShowAuthor then
ItemTemplate=ItemTemplate&"{AUTHOR}"
end if
ItemTemplate=ItemTemplate&"</div>"
end if
if bShowCategory then
ItemTemplate=ItemTemplate&"<div class='QS_feeditemcategory'>{CATEGORY}</div>"
end if
ItemTemplate=ItemTemplate&"<div class='QS_feeditemdescription'>{DESCRIPTION}</div>"
ItemTemplate=ItemTemplate&"</div>"
end if
dim FeedItems,FeedItemsCount, feedItem, child, feedlink, feeddescription, feedDate, feedCommentsLink, CategoryItems, feedAuthor, feedenclosure
dim feedCategory, categoryitem, i, feedtitle, RSGUID, node, feedimage
dim arrLinki,arrLinks
if not isLeeg(sUrls) then
arrLinks=split(sUrls,vbcrlf)
else
redim arrLinks(0)
arrLinks(0)=sLink
end if
for arrLinki=lbound(arrLinks) to ubound(arrLinks) 'start loop
if not isLeeg(arrLinks(arrLinki)) then
if instr(lcase(arrLinks(arrLinki)),"rss.asp?iid=")<>0 then
if convertStr(application("QSRSSLINKLOADED"))<>"1" then
application("QSRSSLINKLOADED")="1"
response.redirect(selectedPage.getSimpleLink)
end if
end if

dim xmlDOM
Set xmlDOM = Server.CreateObject("msxml2.DOMDocument")
xmlDOM.async = false
xmlDOM.setProperty "ServerHTTPRequest", True
xmlDom.resolveExternals = False

dim letsgo
letsgo=false

if xmlDOM.Load(arrLinks(arrLinki)) then
    letsgo=true
else
    If xmlDOM.parseError.errorCode <> 0 Then
        Set xmlDOM = Server.CreateObject("Msxml2.DOMDocument.6.0")
        xmlDOM.async = false
        xmlDOM.setProperty "ServerHTTPRequest", True
        xmlDom.resolveExternals = False
        if xmlDOM.Load(arrLinks(arrLinki)) then letsgo=true
    End If
end if

If letsgo Then 
 
if err.number <> 0 then 
exit function
end if

if xmlDOM.getElementsByTagName("rss").length >0 then 
Set FeedItems = xmlDOM.getElementsByTagName("item")
feedType="RSS"
end if
if xmlDOM.getElementsByTagName("rdf:RDF").length >0 then 
Set FeedItems = xmlDOM.getElementsByTagName("item")
feedType="RSS"
end if
if xmlDOM.getElementsByTagName("feed").length >0 then 
Set FeedItems = xmlDOM.getElementsByTagName("entry")
feedType="ATOM"
end if
if xmlDOM.getElementsByTagName("atom:feed").length >0 then 
Set FeedItems = xmlDOM.getElementsByTagName("atom:entry")
feedType="ATOM"
end if
if not isLeeg(feedType) then
FeedItemsCount = FeedItems.Length-1
j = -1
For i = 0 To FeedItemsCount
feedtitle=""
feedlink=""
feedAuthor=""
feeddescription=""
feedDate=""
feedCommentsLink=""
feedCategory=""
feedenclosure=""
Set feedItem = FeedItems.Item(i)
for each child in feedItem.childNodes
select case feedType
case "RSS"
Select case lcase(child.nodeName)
case "title"
feedtitle = child.text
case "link"
feedlink = child.text
case "author","dc:creator"
feedAuthor=child.text
case "description"
feeddescription = child.text
case "pubdate","dc:date"
feedDate = child.text
case "image"
feedimage = child.text
case "enclosure"
feedenclosure = child.GetAttribute("url")
case "guid"
if isLeeg(feedlink) then feedlink=child.text
case "comments"
feedCommentsLink = child.text
case "category"
Set CategoryItems = feedItem.getElementsByTagName("category")
feedCategory = ""
for each categoryitem in CategoryItems
if feedCategory <> "" Then 
feedCategory = feedCategory & ", "
End If
feedCategory = feedCategory & categoryitem.text
Next
End Select
case "ATOM"
Select case lcase(child.nodeName)
case "updated","atom:updated"
feedDate = left(child.text,10) & " " & mid(child.text,12,5)
case "author","atom:author"
dim authorName, authorEmail, authoruri
for each node in child.childnodes
select case lcase(convertStr(node.nodeName))
case "name","atom:name"
authorName=node.text
case "email","atom:email"
authorEmail=node.text
case "uri","atom:uri"
authoruri=node.text
end select
next
feedAuthor=trim(authorName & " " & authorEmail & " " & authoruri)
Case "link","atom:link"
if lcase(convertStr(child.GetAttribute("rel")))="alternate" or isleeg(child.GetAttribute("rel")) then
feedlink = child.GetAttribute("href")
end if
Case "title","atom:title"
feedtitle = child.text
case "id","atom:id"
RSGUID	= child.text
Case "content","atom:content"
feeddescription = child.text
case "atom:icon"
feedimage = child.text
case "atom:logo"
feedimage = child.text
case "media:thumbnail"
dim att
for each att in child.attributes
if lcase(att.name)="url" then
feedimage = att.text
end if
next
End Select
End Select
next
' now check filter
'If (InStr(feedTitle,Keyword1)>0) or (InStr(feedTitle,Keyword2)>0) or (InStr(feedDescription,Keyword1)>0) or (InStr(feedDescription,Keyword2)>0) then
dim j, ItemContent, iconLink, commentLink
j = J+1
if J<iMaxItems or convertBool(bRandom) then 
ItemContent = Replace(ItemTemplate,"{LINK}",feedlink,1,-1,1)

if convertGetal(iTitleLimitTo)<>0 then	
		
	feedTitle=removeHTML(feedTitle)
	if len(feedTitle)>convertGetal(iTitleLimitTo) then
		feedTitle=left(feedTitle,iTitleLimitTo) & " ..."
	end if
	
end if


ItemContent = Replace(ItemContent,"{TITLE}",feedTitle,1,-1,1)
ItemContent = Replace(ItemContent,"{AUTHOR}",linkUrls(feedAuthor),1,-1,1)
ItemContent = Replace(ItemContent,"{DATE}",feedDate,1,-1,1)
ItemContent = Replace(ItemContent,"{CATEGORY}",feedCategory,1,-1,1)
ItemContent = Replace(ItemContent,"{IMAGE}",feedimage,1,-1,1)
ItemContent = Replace(ItemContent,"{ENCLOSURE}",feedenclosure,1,-1,1)
ItemContent = Replace(ItemContent,"{COUNTER}",j+1,1,-1,1)
if bOpenLinkInNW then
iconLink	=	commentLink
end if
if convertGetal(iLimitTo)<>0 then
	if convertGetal(iLimitTo)=QS_feedNoText then
		feedDescription=""
	else
		feedDescription=removeHTML(feedDescription)
		if len(feedDescription)>convertGetal(iLimitTo) then
			feedDescription=left(feedDescription,iLimitTo) & " ..."
		end if
	end if
end if
if bOpenLinkInNW then
ItemContent=Replace(ItemContent,"{DESCRIPTION}",openLinksInNW(feedDescription),1,-1,1)
else
ItemContent=Replace(ItemContent,"{DESCRIPTION}",feedDescription,1,-1,1)
end if
build=build & ItemContent & "##-QS_DELIMITER-##"
ItemContent = ""
iconLink	= ""
feedCommentsLink	= ""
End if
'End If 
Next
if not bTemplate then 
'lege items eruit halen	- alleen in het geval geen template werd gebruikt
build=replace(build,"<div class='QS_feeditemcategory'></div>","",1,-1,1)
build=replace(build,"<div class='QS_feeditemdescription'></div>","",1,-1,1)
build=replace(build,"<div class='QS_feeditemdetails'></div>","",1,-1,1)
build=replace(build,"&nbsp;|&nbsp;</div>","</div>",1,-1,1)
else
build=treatConstants(build,false)
end if
end if
end if
end if
next 'end loop!
Set xmlDOM = Nothing
application(QS_CMS_cacheFEED & iId)=build
application(QS_CMS_cacheFEED & iId & "TS")=dateAdd("s",iCache,now())
'dumperror "feed: "& sName, err
On Error Goto 0
else
build=application(QS_CMS_cacheFEED & iId)
end if
if bRandom then
dim arrF
arrF=split(build,"##-QS_DELIMITER-##")
dim DictItems
set DictItems=server.CreateObject ("scripting.dictionary")
dim fRn
for fRn=lbound(arrF) to ubound(arrF)
if not isLeeg(arrF(fRn)) then
DictItems.Add  generatepassword & fRn, arrF(fRn)
end if
next
SortDictionary2 DictItems,1
build=""
dim fItem,runnerI
runnerI=0
for each fItem in DictItems
build=build & DictItems(fItem)
runnerI=runnerI+1
if runnerI=iMaxItems then exit for
next
set DictItems=nothing
else
build=replace(build,"##-QS_DELIMITER-##","",1,-1,1)
end if
if not bTemplate then
if convertBool(bEnableJS) then
build="<div id='QS_feed'>" & treatConstants(build,false) & "</div>"
else
build="<div id='QS_feed'>" & removeJS(treatConstants(build,false)) & "</div>"
end if
build=build& "<div style=""float:none;clear:both;margin:0;padding:0;border:none;font-size: 1px;""></div>"
else
'de html er voor en achter zetten
if convertBool(bEnableJS) then
build=treatConstants(sHTMLBefore,true) & build & treatConstants(sHTMLAfter,true)
else
build=treatConstants(sHTMLBefore,true) & removeJS(build) & treatConstants(sHTMLAfter,true)
end if
end if
end if
end function
public function copyToCustomer(oCID)
if isNumeriek(iId) then
dim oldId
oldID=iId
overruleCID	= oCID
iId=null
save()
db.execute("update tblPage set iFeedID=" & iId & " where iFeedID=" & oldID & " and iCustomerID=" & oCID)
end if
end function
end class%>
