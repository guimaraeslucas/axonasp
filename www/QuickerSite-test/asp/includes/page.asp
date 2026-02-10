
<%dim allSubPages
set allSubPages=server.CreateObject ("scripting.dictionary")
class cls_page
Public iId, iParentID, iListPageID, iCustomerID, sTitle, sExternalURLPrefix, sExternalURL, bOpenInNewWindow,sValueToBeValidated,sTitleToBeValidated
Public sValue, sValueTextOnly, iRang, bOnline, bDeleted, bContainerPage, bLossePagina, bHomepage, updatedTS,iUpdatedBy,dUpdatedOn
Public createdTS, sPw, dOnlineFrom, dOnlineUntill, sOrderBY, dPage, iHits, iVisitors, sApplication, sLPExternalURL,bHideDate
Public sCode, bOpenOnload, bIntranet, iCatalogId, iFormID, sFormAlign, bMenuGroup, iFeedId, sKeywords, sSEOtitle, bCopy
Public sDescription, iTemplateID, bPushRSS, iHitsRSS, sRSSLink, sUserFriendlyURL, sRedirectTo, iReload, iThemeID, sRel,sLPIC
Public bLPExternalOINW, newRang, oldRang, sHeader, iLPOpenByDefault, overruleCID, bam,p_template, iPostID, forceArtisteerRoot, sPageTitle, b404, sUfl
Public sProp01,sProp02,sProp03,sProp04,sProp05,sProp06,sProp07,sProp08,sClassname,sPageCache,bNocache,sUrlRRSImage,bAccordeon,bHideFromSearch,sItemPicture

Public iPMlocation

Private bLoadInCache
Private Sub Class_Initialize
on error resume next
bNocache=false
bLoadInCache=false
iId=null
iCustomerID=cId
iParentID=convertLng(decrypt(request("iParentID")))
iListPageID=convertLng(decrypt(request("iListPageID")))
bLossePagina=convertBool(Request("bLossePagina"))
bIntranet=request("bIntranet")
bOnline=true
bDeleted=false
bContainerPage=false
bOpenOnload=false
bOpenInNewWindow=true
bHomepage=0
dOnlineUntill=null
dOnlineFrom=null
iCatalogId=null
iFeedId=null
iFormID=null
bMenuGroup=false
bPushRSS=false
iReload=0
iThemeID=null
sLPExternalURL="http://"
bLPExternalOINW=false
newRang=0
iLPOpenByDefault=0
overruleCID=null
bam=false
iUpdatedBy=null
forceArtisteerRoot=false
bHideDate=false
bCopy=false
bAccordeon=false
bHideFromSearch=false
iPMlocation=0
b404=false
set p_template=nothing
on error goto 0
end sub
Public function pickByCode(sPageCode)
ON Error Resume Next
if not isLeeg(sPageCode) then
dim rs
set rs=db.execute("select iID from tblPage where " & sqlCustId & " and sCode='"& ucase(cleanup(sPageCode)) & "'")
if not rs.eof then
pick(rs(0))
end if
set rs=nothing
dumperror "Pick Page By Code",err
end if
ON Error Goto 0
end function
Public function pickByCodeNOCID(sPageCode)
ON Error Resume Next
if not isLeeg(sPageCode) then
dim rs
set rs=db.execute("select iID from tblPage where sCode='"& ucase(cleanup(sPageCode)) & "'")
if not rs.eof then
pick(rs(0))
end if
set rs=nothing
dumperror "Pick Page By Code",err
end if
ON Error Goto 0
end function
private sub forceUFL()
On Error Resume Next
if isLeeg(iId) and not bCopy then 'new records only (not copy)
if isleeg(sUserFriendlyURL) then
if convertBool(customer.bUserFriendlyURL) then
sUserFriendlyURL=replace(sTitle," ","-",1,-1,1)
sUserFriendlyURL=replace(sUserFriendlyURL,"?","",1,-1,1)
sUserFriendlyURL=replace(sUserFriendlyURL,"!","",1,-1,1)
sUserFriendlyURL=replace(sUserFriendlyURL,"'","",1,-1,1)
sUserFriendlyURL=replace(sUserFriendlyURL,"/","-",1,-1,1)
sUserFriendlyURL=replace(sUserFriendlyURL,"\","-",1,-1,1)
sUserFriendlyURL=replace(sUserFriendlyURL,"--","-",1,-1,1)
sUserFriendlyURL=replace(sUserFriendlyURL,"--","-",1,-1,1)
sUserFriendlyURL=replace(sUserFriendlyURL,"--","-",1,-1,1)
sUserFriendlyURL=replace(sUserFriendlyURL,"--","-",1,-1,1)
sUserFriendlyURL=replace(sUserFriendlyURL,"--","-",1,-1,1)
sUserFriendlyURL=replace(sUserFriendlyURL,"--","-",1,-1,1)
sUserFriendlyURL=replace(sUserFriendlyURL,"--","-",1,-1,1)
if right(sUserFriendlyURL,1)="-" then
sUserFriendlyURL=left(sUserFriendlyURL,len(sUserFriendlyURL)-1)
end if
if IsAlphaNumeric(sUserFriendlyURL) then
sUserFriendlyURL=left(lcase(sUserFriendlyURL),50)
else
sUserFriendlyURL=""
end if
end if
end if
end if
on error goto 0
end sub
Public sub pickByUserFriendlyURL
dim s404ufl
if not convertBool(customer.bUserFriendlyURL) then
exit sub
end if
dim ufl,helpufl
ufl=Request.ServerVariables("query_string")
if instr(ufl,"?")>0 then
ufl=left(ufl,instr(ufl,"?"))
end if
helpufl="QS_UFL_" & ufl
if convertGetal(application(helpufl))<>0 then
on error resume next
pick(application(helpufl))
Response.Status = "200 OK"
on error goto 0
showFromCache()
exit sub
end if
if instr(lcase(ufl),"pageaction=")<>0 then
exit sub
end if
select case right(lcase(ufl),4)
case ".ico",".gif",".jpg",".png",".css",".map"
response.end
end select
if not isLeeg(ufl) then
if not isLeeg(customer.sAlternateDomains) then
dim alternatedomains,ad
alternatedomains=split(customer.sAlternateDomains,vbcrlf)
for ad=lbound(alternatedomains) to ubound(alternatedomains)
ufl=replace(ufl,alternatedomains(ad),"",1,-1,1)
next
end if
ufl=replace(ufl,":80","",1,-1,1)
ufl=replace(ufl,":443","",1,-1,1)
ufl=replace(ufl,"http://" & request.servervariables("http_host") & C_VIRT_DIR,"",1,-1,1)
ufl=replace(ufl,"https://" & request.servervariables("http_host") & C_VIRT_DIR,"",1,-1,1) 
ufl=replace(ufl,"404;","",1,-1,1)
ufl=replace(ufl,"aspxerrorpath=/","",1,-1,1)
ufl=replace(ufl,customer.sUrl,"",1,-1,1)
ufl=replace(ufl,":"&request.servervariables("server_port"),"",1,-1,1)

s404ufl=ufl

dim bQSuseVD,sQSuseVDPath
bQSuseVD=false
if C_DIRECTORY_QUICKERSITE<>"" then
dim lcq
lcq=len(C_DIRECTORY_QUICKERSITE)+1
if left(ufl,lcq)=C_DIRECTORY_QUICKERSITE&"/" then
ufl=right(ufl,len(ufl)-lcq)
else
'NEW - was response.redirect(C_DIRECTORY_QUICKERSITE & ufl)
if lcase(convertStr(ufl))="/backsite" then response.redirect(C_DIRECTORY_QUICKERSITE & ufl)
bQSuseVD=true
sQSuseVDPath=C_DIRECTORY_QUICKERSITE & ufl
end if
end if
dim uflLoop
uflLoop=true
while uflLoop
if instr(ufl,"//")>0 then
ufl=replace(ufl,"//","/",1,-1,1)
else
uflLoop=false
end if
wend
uflLoop=true
while uflLoop
if left(ufl,1)="/" then
ufl=right(ufl,len(ufl)-1)
else
uflLoop=false
end if
wend
uflLoop=true
while uflLoop
if right(ufl,1)="/" then
ufl=left(ufl,len(ufl)-1)
Response.Redirect ("/"&ufl)
else
uflLoop=false
end if
wend
ufl=left(ufl,200)
ON Error Resume Next
if not isLeeg(ufl) and IsAlphaNumeric(ufl) then
dim rs
set rs=db.execute("select iID from tblPage where bDeleted="&getSQLBoolean(false)&" and " & sqlCustId & " and sUserFriendlyURL='"& ufl & "'")
if not rs.eof then
'new! if the function showFromCache will define whether or not a cached copy of the page should be shown
'so a cached copy of a page will only be shown in case the visitor browse the a page with a userfriendly url, never with a .asp extension
pick(rs(0))
Response.Status = "200 OK"
showFromCache()
'new
if bQSuseVD then
if isLeeg(sApplication) then
response.redirect(sQSuseVDPath)
end if
end if
if instr(ufl,"/")<>0 then Response.Redirect (C_DIRECTORY_QUICKERSITE&"/default.asp?iId="& EnCrypt(iId))
application(helpufl)=iId
else

'custom error page?
if convertBool(customer.bCustom404) then

	b404=true : sUfl=ufl

else

	redirectToHP()

end if


end if
set rs=nothing
end if
ON Error Goto 0
end if
end sub
private property get sTitleForMenu
if bam and (convertGetal(iParentID)=0 or forceArtisteerRoot) and not bIntranet then
sTitleForMenu=getParentLinkForArtisteer(quotrep(sTitle))
else
sTitleForMenu=quotrep(sTitle)
end if
end property
Public Function Pick(id)
ON Error Resume Next
dim sql, RS
if isNumeriek(id) then
'Response.Write "test: " & id & "<br />"
if isNull(overruleCID) then
sql = "select * from tblPage where "& sqlCustId & " and iId=" & id
else
sql = "select * from tblPage where iId=" & id
end if
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
iParentID	= rs("iParentID")
iListPageID	= rs("iListPageID")
iCustomerID	= rs("iCustomerID")
sTitle	= rs("sTitle")
sValue	= rs("sValue")
sValueTextOnly	= rs("sValueTextOnly")
sExternalURLPrefix	= rs("sExternalURLPrefix")
sExternalURL	= rs("sExternalURL")
bOpenInNewWindow	= rs("bOpenInNewWindow")
iRang	= rs("iRang")
bOnline	= rs("bOnline")
bDeleted	= rs("bDeleted")
bHomepage	= rs("bHomepage")
bContainerPage	= rs("bContainerPage")
bLossePagina	= rs("bLossePagina")
sPw	= rs("sPw")
updatedTS	= rs("updatedTS")
createdTS	= rs("createdTS")
dOnlineUntill	= rs("dOnlineUntill")
dOnlineFrom	= rs("dOnlineFrom")
sOrderBY	= rs("sOrderBY")
dPage	= rs("dPage")
iHits	= convertGetal(rs("iHits"))
iVisitors	= convertGetal(rs("iVisitors"))
sApplication	= rs("sApplication")
sCode	= rs("sCode")
bOpenOnload	= rs("bOpenOnload")
bIntranet	= rs("bIntranet")
iCatalogId	= rs("iCatalogId")
iFormID	= rs("iFormID")
sFormAlign	= rs("sFormAlign")
bMenuGroup	= rs("bMenuGroup")
iFeedId	= rs("iFeedId")
sKeywords	= rs("sKeywords")
sDescription	= rs("sDescription")
iTemplateID	= rs("iTemplateID")
bPushRSS	= rs("bPushRSS")
iHitsRSS	= rs("iHitsRSS")
sRSSLink	= rs("sRSSLink")
sUserFriendlyURL	= rs("sUserFriendlyURL")
iReload	= rs("iReload")
sRedirectTo	= rs("sRedirectTo")
iThemeID	= rs("iThemeID")
sSEOtitle	= rs("sSEOtitle")
sLPExternalURL	= rs("sLPExternalURL")
bLPExternalOINW	= rs("bLPExternalOINW")
sHeader	= rs("sHeader")
iLPOpenByDefault	= rs("iLPOpenByDefault")
sPageTitle	= rs("sPageTitle")
sValueToBeValidated	= rs("sValueToBeValidated")
sTitleToBeValidated	= rs("sTitleToBeValidated")
iUpdatedBy	= rs("iUpdatedBy")
dUpdatedOn	= rs("dUpdatedOn")
sRel	= rs("sRel")
bHideDate	= rs("bHideDate")
sClassname	= rs("sClassname")
sProp01	= rs("sProp01")
sProp02	= rs("sProp02")
sProp03	= rs("sProp03")
sProp04	= rs("sProp04")
sProp05	= rs("sProp05")
sProp06	= rs("sProp06")
sProp07	= rs("sProp07")
sProp08	= rs("sProp08")
sPageCache	= rs("sPageCache")
bNocache	= rs("bNocache")
bAccordeon	= rs("bAccordeon")
bHideFromSearch	= rs("bHideFromSearch")
sItemPicture	= rs("sItemPicture")
sLPIC=rs("sLPIC")
iPMlocation=convertGetal(rs("iPMlocation"))
'sUrlRRSImage	= rs("sUrlRRSImage")
dumperror "Pick Page",err
end if
set RS = nothing
end if
ON Error Goto 0
end function
Public Function Check()
Check = true
dim rs
if bDeleted then
exit function
end if
if isLeeg(sTitle) then
check=false
message.AddError("err_mandatory")
end if
if lcase(convertStr(sLPExternalURL))="http://" then
sLPExternalURL=""
end if
sValue=removeEmptyP(sValue)
if sValue="<br class=""innova"">" then sValue=""
if isLeeg(sValue) then
if isLeeg(sExternalURL) and isLeeg(sLPExternalURL) then
if not bMenuGroup and not bContainerPage and isLeeg(sApplication) and isLeeg(iCatalogId) and isLeeg(iFormID) and isLeeg(iFeedId) and isLeeg(iThemeID) and isLeeg(sOrderBY) then
check=false
message.AddError("err_mandatory")
end if
if bHomepage and isLeeg(sApplication)  and isLeeg(iCatalogId) and isLeeg(iFormID) and isLeeg(iFeedId) and isLeeg(iThemeID) and isLeeg(sOrderBY) then
message.AddError("err_mandatory")
check=false
end if
end if 
end if
if not isLeeg(sCode)	then
set rs=db.execute("select iId from tblPage where bDeleted="&getSQLBoolean(false)&" and iId<>"&  convertGetal(iId) & " and " &  sqlCustId & " and sCode='" & ucase(sCode) & "'")
if not rs.eof then
check=false
message.AddError("err_appCode")
end if
set rs=nothing
end if
if not isLeeg(sLPExternalURL) and not isLeeg(sValue) then
check=false
message.AddError("err_listpagepurl")
end if
if not isLeeg(sLPExternalURL) then
if lcase(convertStr(left(sLPExternalURL,3)))="www" then
sLPExternalURL="http://" & sLPExternalURL
end if
end if
if not isLeeg(removeEmptyP(sValue)) and not convertBool(bContainerPage) then
forceUFL()
end if
if not isLeeg(sUserFriendlyURL) then
sUserFriendlyURL=trim(sUserFriendlyURL)
dim uflLoop
uflLoop=true
while uflLoop
if instr(sUserFriendlyURL,"//")<>0 then
sUserFriendlyURL=replace(sUserFriendlyURL,"//","/",1,-1,1)
else
uflLoop=false
end if
wend
if left(sUserFriendlyURL,1)="/" then
sUserFriendlyURL=right(sUserFriendlyURL,len(sUserFriendlyURL)-1)
end if
if right(sUserFriendlyURL,1)="/" then
sUserFriendlyURL=left(sUserFriendlyURL,len(sUserFriendlyURL)-1)
end if
if IsAlphaNumeric(sUserFriendlyURL) then
set rs=db.execute("select iId from tblPage where bDeleted="&getSQLBoolean(false)&" and iId<>"&  convertGetal(iId) & " and " &  sqlCustId & " and sUserFriendlyURL='" & sUserFriendlyURL & "'")
if not rs.eof then
check=false
message.AddError("err_ufl")
end if
set rs=nothing
else
check=false
message.AddError("err_ufl_an")
end if
end if
if not isLeeg(sExternalURLPrefix)  then
if isLeeg(sExternalURL) then
message.AddError("err_mandatory")
check=false
end if
end if
if not isLeeg(iCatalogId) and not isLeeg(iFormID) then
message.AddError("err_catorform")
check=false
end if
End Function
Public Function Save
On Error Resume Next
if isNull(overruleCID) then
if check() then
save=true
else
save=false
exit function
end if
end if
if bHomepage then 
bOnline=true
bContainerPage=false
end if
if isLeeg(sPw) and not isLeeg(parentPage.sPw) then
sPw=parentPage.sPw
end if
if isLeeg(sPw) and not isLeeg(customer.sTotalPW) then
sPw=customer.sTotalPW
end if
if convertBool(parentPage.bIntranet) then
bIntranet=true
end if
if isLeeg(bIntranet) then
bIntranet=false
end if
if isLeeg(sTitleToBeValidated) then
sTitleToBeValidated=null
end if
if isLeeg(sValueToBeValidated) then
sValueToBeValidated=null
end if
if isLeeg(bNocache) then bNocache=false
'check URLS
sExternalURL=convertStr(sExternalURL)
sExternalURL=replace(sExternalURL,"http://","",1,-1,1)
sExternalURL=replace(sExternalURL,"https://","",1,-1,1)
sExternalURL=replace(sExternalURL,"ftp://","",1,-1,1)
sExternalURL=replace(sExternalURL,"mailto:","",1,-1,1)
'set db=nothing
'set db=new cls_database
dim rs
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblPage where 1=2"
rs.AddNew
rs("createdTS")=now()
if isNull(overruleCID) then
iRang=getRang+1
end if
else
if isNull(overruleCID) then
rs.Open "select * from tblPage where "& sqlCustId &" and iId="& iId
else
rs.Open "select * from tblPage where iId="& iId
end if
end if
rs("iParentid")	= iParentID
rs("iListPageID")	= iListPageID
rs("iCustomerID")	= iCustomerID
rs("sTitle")	= left(sTitle,250)
rs("sExternalURLPrefix")	= sExternalURLPrefix
rs("sExternalURL")	= sExternalURL
rs("sValue")	= sValue
rs("sHeader")	= sHeader
rs("iLPOpenByDefault")	= iLPOpenByDefault
rs("sValueToBeValidated")	= sValueToBeValidated
rs("iUpdatedBy")	= iUpdatedBy
rs("bHideDate")	= convertBool(bHideDate)
rs("sClassname")	= sClassname
dim cValue
cValue=sValue
rs("sValueTextOnly")	= RemoveHTML(treatConstants(cValue,false))
dim setRangByUser
setRangByUser=false
if isLeeg(iListPageID) and not bLossePagina then
if convertGetal(newRang)<>0 and convertGetal(newRang)<>convertGetal(iRang) then
if secondAdmin.bPagesMove then
'set new rang set by user!
oldRang	= iRang
iRang	= newRang
rs("iRang")	= iRang
setRangByUser	= true
end if
else
if convertGetal(iRang)=0 then
rs("iRang")	= getRang+1
else
rs("iRang")	= iRang
end if
end if
else
rs("iRang")	= 0
end if
rs("bOnline")	= bOnline
rs("bHomepage")	= bHomepage
rs("bDeleted")	= bDeleted
rs("bContainerPage")	= bContainerPage
rs("bOpenInNewWindow")	= bOpenInNewWindow
rs("bLossePagina")	= bLossePagina
rs("updatedTS")	= now()
rs("sPw")	= sPw
rs("dOnlineFrom")	= dOnlineFrom
rs("dOnlineUntill")	= dOnlineUntill
rs("sOrderBY")	= sOrderBY
rs("dPage")	= dPage
rs("sApplication")	= sApplication
rs("sCode")	= sCode
rs("bOpenOnload")	= bOpenOnload
rs("bIntranet")	= bIntranet
rs("iCatalogId")	= iCatalogId
rs("iFormID")	= iFormID
rs("sFormAlign")	= sFormAlign
rs("iFeedId")	= iFeedId
rs("sKeywords")	= sKeywords
rs("sDescription")	= sDescription
rs("iTemplateID")	= iTemplateID
rs("bPushRSS")	= bPushRSS
rs("iHitsRSS")	= iHitsRSS
rs("sRSSLink")	= sRSSLink
rs("sUserFriendlyURL")	= sUserFriendlyURL
rs("sRedirectTo")	= sRedirectTo
rs("iReload")	= iReload
rs("iThemeID")	= iThemeID
rs("sSEOtitle")	= sSEOtitle
rs("sLPExternalURL")	= sLPExternalURL
rs("bLPExternalOINW")	= bLPExternalOINW
rs("sPageTitle")	= sPageTitle
rs("bMenuGroup")	= false
rs("sTitleToBeValidated") = sTitleToBeValidated
rs("dUpdatedOn")	= dUpdatedOn
rs("sRel")	= sRel
rs("sProp01")	= sProp01
rs("sProp02")	= sProp02
rs("sProp03")	= sProp03
rs("sProp04")	= sProp04
rs("sProp05")	= sProp05
rs("sProp06")	= sProp06
rs("sProp07")	= sProp07
rs("sProp08")	= sProp08
rs("bNocache")	= bNocache
rs("bAccordeon")	= convertBool(bAccordeon)
rs("bHideFromSearch")	= convertBool(bHideFromSearch)
rs("sItemPicture") = convertStr(sItemPicture)
rs("sLPIC") = sLPIC
rs("iPMlocation")=iPMlocation

'rs("sUrlRRSImage")	= sUrlRRSImage
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
if bHomepage then
dim sql
sql="update tblPage set bHomepage="&getSQLBoolean(false)&" where iCustomerId="& iCustomerID &" and iId<>"& iId
if bIntranet then
sql=sql&" and bIntranet="&getSQLBoolean(true)&" "
else
sql=sql&" and (bIntranet="&getSQLBoolean(false)&" or bIntranet is null) "
end if
db.execute(sql)
end if
if not bOnline then
getAllSubPagesOffline(iId)
end if
if bDeleted then
db.execute("delete from tblContactPage where iTitleID=" & iId & " or iBodyID=" & iId)
end if
if setRangByUser then
setRangByNumber oldRang,newRang
end if
if not convertbool(bLossePagina) and convertGetal(iListPageID)=0 then
clearMenuCache
end if
clearRSScache
Application("iHomePageID")=0
Application("iIntranetHomePageID")=0
dim appitem
for each appitem in application.contents
if left(convertStr(appitem),7)="QS_UFL_" then
application.contents.remove(appitem)
end if
next
if bDeleted=true then
deleteListItemImage
end if
On Error Goto 0
end function
public sub clearRSScache
application("RSS"&cId&iListPageID)	= ""
application("RSS"&cId&iId)	= ""
end sub
public function getRequestValues()
'secured!
if secondAdmin.bPageTitle then 
sTitle = convertStr(Request.Form ("sTitle"))
sPageTitle = convertStr(Request.Form("sPageTitle"))
end if
if secondAdmin.bPageOrder then newRang	= convertGetal(Request.Form ("iRang"))
if secondAdmin.bPagePublish then bOnline = convertBool(Request.Form ("bOnline"))
if secondAdmin.bPageSetHomepage then bHomepage = convertBool(Request.Form ("bHomepage"))
if secondAdmin.bPageFeed then iFeedId = convertLng(decrypt(Request.Form ("iFeedId")))
if secondAdmin.bPageCatalog then iCatalogId	= convertLng(decrypt(Request.Form ("iCatalogId")))
if secondAdmin.bPageTemplate then iTemplateID = convertLng(decrypt(Request.Form ("iTemplateID")))
if secondAdmin.bPageForm then iFormID = convertLng(decrypt(Request.Form ("iFormID")))
if secondAdmin.bPageForm then sFormAlign = convertStr(Request.Form ("sFormAlign"))
if secondAdmin.bApplicationpath then sApplication = convertStr(Request.Form ("sApplication"))
if secondAdmin.bApplicationpath then sCode = ucase(convertStr(Request.Form ("sCode")))
if secondAdmin.bPageTheme then iThemeID	= convertLng(decrypt(Request.Form ("iThemeID")))
if secondAdmin.bPageRefresh then sRedirectTo = convertStr(Request.Form ("sRedirectTo"))
if secondAdmin.bPageRefresh then iReload = convertGetal(Request.Form ("iReload"))
if secondAdmin.bPageAdditionalHeader then sHeader = convertStr(Request.Form ("sHeader"))
if secondAdmin.bPageKeywords then sKeywords	= convertStr(Request.Form ("sKeywords"))
if secondAdmin.bPageDescription then sDescription = convertStr(Request.Form ("sDescription"))
if secondAdmin.bPageUrlRSS then sRSSLink = convertStr(Request.Form ("sRSSLink"))
if secondAdmin.bPageTitleTag then sSEOtitle = convertStr(Request.Form ("sSEOtitle"))
if secondAdmin.bPageUFL then sUserFriendlyURL = replace(convertStr(Request.Form ("sUserFriendlyURL"))," ","-",1,-1,1)
if secondAdmin.bPageBody then sValue = convertStr(Request.Form ("sValue"))
if secondAdmin.bPageBody then sExternalURLPrefix	= convertStr(Request.Form ("sExternalURLPrefix"))
if secondAdmin.bPageBody then sExternalURL	= convertStr(Request.Form ("sExternalURL"))
if secondAdmin.bPageBody then bOpenInNewWindow	= convertBool(Request.Form ("bOpenInNewWindow"))
if secondAdmin.bPageBody then iLPOpenByDefault = convertGetal(Request.Form ("iLPOpenByDefault"))
iParentID	= convertLng(decrypt(Request.Form ("iParentID")))
iListPageID	= convertLng(decrypt(Request.Form ("iListPageID")))
bContainerPage	= convertBool(Request.Form ("bContainerPage"))
bLossePagina	= convertBool(Request.Form ("bLossePagina"))
dOnlineFrom	= convertDateFromPicker(Request.Form ("dOnlineFrom"))
dOnlineUntill	= convertDateFromPicker(Request.Form ("dOnlineUntill"))
'sUrlRRSImage 	= convertStr(Request.Form ("sUrlRRSImage")) 
sOrderBY	= convertStr(Request.Form ("sOrderBY"))
dPage	= convertDateFromPicker(Request.Form ("dPage"))
bOpenOnload	= convertBool(Request.Form ("bOpenOnload"))
sLPExternalURL	= convertStr(Request.Form("sLPExternalURL"))
sClassname	= trim(convertStr(Request.Form("sClassname")))
bLPExternalOINW	= convertBool(Request.Form ("bLPExternalOINW"))
bPushRSS	= convertBool(Request.Form ("bPushRSS"))
bIntranet	= convertBool(Request.Form ("bIntranet"))
bMenuGroup	= convertBool(Request.Form ("bMenuGroup"))
sRel	= trim(convertStr(request.form("sRel")))
bHideDate	= convertBool(Request.Form ("bHideDate"))
bNocache	= convertBool(request.form("bNocache"))
bAccordeon	= convertBool(request.form("bAccordeon"))
bHideFromSearch	= convertBool(request.form("bHideFromSearch"))
end function
public function subPages(onlineOnly)
set subPages=server.CreateObject ("scripting.dictionary")
if not isLeeg(iId) then
dim rs,sql
sql="select iId from tblPage where "& sqlCustId &" and bDeleted="&getSQLBoolean(false)&" "
if onlineOnly then
sql=sql&" and bOnline="&getSQLBoolean(true)
end if
sql=sql&" and "& sqlNull("iParentID",iId)
sql=sql&" order by iRang asc"
set rs=db.execute(sql)
dim subPage
while not rs.eof
set subPage=new cls_page
subPage.Pick (rs(0))
subPages.Add subPage.iId, subPage
set subPage=nothing
rs.movenext
wend
set rs=nothing
end if
end function
public function canBeDeleted()
if not secondAdmin.bPagesDelete then
canBeDeleted=false
else
dim copySubPages
set copySubPages=subPages(false)
if copySubPages.count=0 and not bHomepage and not isLeeg(iId) then
canBeDeleted=true
else
canBeDeleted=false
end if
end if
end function
public function canGoOffline()
if convertGetal(iId)=convertGetal(getHomePage.iId) or convertGetal(iId)=convertGetal(getIntranetHomePage.iId) then
canGoOffline=false
exit function
end if
getAllSubPages(iId)
if allSubPages.Exists (getHomePage.iId) or allSubPages.Exists (getIntranetHomePage.iId) then
canGoOffline=false
else
canGoOffline=true
end if
end function
public function parentPage
set parentPage=new cls_page
if not isLeeg(iParentID) then
parentPage.Pick (iParentID)
end if
end function
public function listPage
set listPage=new cls_page
if not isLeeg(iListPageID) then
listPage.Pick (iListPageID)
end if
end function
public function getAllSubPagesOffline (pageId)
dim page
set page=new cls_page
page.pick(pageId)
if isLeeg(page.iId) then exit function
dim copySubPages, subPage
set copySubPages=page.subPages(false)
for each subPage in copySubPages
copySubPages(subPage).bOnline=false
copySubPages(subPage).save()
getAllSubPagesOffline(copySubPages(subPage).iId)
next
set copySubPages=nothing
clearMenuCache
end function
public function getAllSubPages(pageId)
On Error Resume Next
dim page
set page=new cls_page
page.pick(pageId)
if isLeeg(page.iId) then exit function
dim copySubPages, subPage
set copySubPages=page.subPages(false)
for each subPage in copySubPages
allSubPages.Add subPage, ""
getAllSubPages(subPage)
next
set copySubPages=nothing
On Error Goto 0
end function
public property get hasSiblings
hasSiblings=false
dim rs,sql
sql="select iId from tblPage where "& sqlCustId &" and bDeleted="&getSQLBoolean(false)&" and (bLossePagina="&getSQLBoolean(false)&" or bLossePagina is null) and iListPageID is null and iId<>"& iId & " and " & sqlNull("iParentID",iParentID)
if convertBool(bIntranet) then
sql=sql& " and bIntranet="&getSQLBoolean(true)&" "
else
sql=sql& " and (bIntranet="&getSQLBoolean(false)&" or bIntranet is null) "
end if
set rs=db.execute(sql)
if not rs.eof then
hasSiblings=true
end if
set rs=nothing
end property
public property get siblings
dim rs,sql
sql="select count(*) from tblPage where "& sqlCustId &" and (bLossePagina="&getSQLBoolean(false)&" or bLossePagina is null) and iListPageID is null and bDeleted="&getSQLBoolean(false)&" and " & sqlNull("iParentID",iParentID)
if convertBool(bIntranet) then
sql=sql& " and bIntranet="&getSQLBoolean(true)&" "
else
sql=sql& " and (bIntranet="&getSQLBoolean(false)&" or bIntranet is null) "
end if
set rs=db.execute(sql)
siblings=clng(rs(0))
set rs=nothing
end property
public property get onlineSiblings
set onlineSiblings=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs,sql
sql="select iId from tblPage where "& sqlCustId &" and iId<>"& iId &" and (bLossePagina="&getSQLBoolean(false)&" or bLossePagina is null) and iListPageID is null and bOnline="&getSQLBoolean(true)&" and bDeleted="&getSQLBoolean(false)&" and " & sqlNull("iParentID",iParentID)
if bIntranet then
sql=sql& " and bIntranet="&getSQLBoolean(true)&" "
else
sql=sql& " and (bIntranet="&getSQLBoolean(false)&" or bIntranet is null) "
end if
sql=sql& " order by iRang desc"
set rs=db.execute(sql)
dim page
while not rs.eof
set page=new cls_page
page.pick(rs(0))
onlineSiblings.Add page.iId,page
set page=nothing
rs.movenext
wend 
set rs=nothing
end if
end property
private function setRangByNumber(oldR,newR)
if not isLeeg(iId) and convertGetal(oldR)<>convertGetal(newR) then
dim sql,rs
if oldR>newR then
sql="update tblPage set iRang=iRang+1 where "
sql=sql&" iRang>=" & newR & " and iRang<" & oldR
else
sql="update tblPage set iRang=iRang-1 where "
sql=sql&" iRang>" & oldR & " and iRang<=" & newR
end if
sql=sql&" and " & sqlCustId &" and (bLossePagina="&getSQLBoolean(false)&" or bLossePagina is null) and iListPageID is null and " & sqlNull("iParentID",iParentID) &" and iId<>"& iId
if convertBool(bIntranet) then
sql=sql& " and bIntranet="&getSQLBoolean(true)&" "
else
sql=sql& " and (bIntranet="&getSQLBoolean(false)&" or bIntranet is null) "
end if
set rs=db.execute(sql)
set rs=nothing
end if
end function
public function moveUp
if not isLeeg(iId) then
if convertGetal(iRang)=1 then
exit function
else
db.execute("update tblPage set iRang=iRang-1 where "& sqlCustId &" and iId="& iId)
dim sql
sql="update tblPage set iRang=iRang+1 where "& sqlCustId &" and (bLossePagina="&getSQLBoolean(false)&" or bLossePagina is null) and iListPageID is null and " & sqlNull("iParentID",iParentID) &" and iId<>"& iId & " and iRang=" & iRang-1
if convertBool(bIntranet) then
sql=sql& " and bIntranet="&getSQLBoolean(true)&" "
else
sql=sql& " and (bIntranet="&getSQLBoolean(false)&" or bIntranet is null) "
end if
db.execute(sql)
end if
clearMenuCache
end if
end function
public function moveDown
if not isLeeg(iId) then
if convertGetal(iRang)=convertGetal(siblings) then
exit function
else
db.execute("update tblPage set iRang=iRang+1 where "& sqlCustId &" and iId="& iId)
dim sql
sql="update tblPage set iRang=iRang-1 where "& sqlCustId &" and (bLossePagina="&getSQLBoolean(false)&" or bLossePagina is null) and iListPageID is null and " & sqlNull("iParentID",iParentID) & " and iId<>"& iId & " and iRang=" & iRang+1
if convertBool(bIntranet) then
sql=sql& " and bIntranet="&getSQLBoolean(true)&" "
else
sql=sql& " and (bIntranet="&getSQLBoolean(false)&" or bIntranet is null) "
end if
db.execute(sql)
end if
clearMenuCache
end if
end function
public function getRang
dim rs,sql
sql="select count(*) from tblPage where "
sql=sql& " tblPage.iCustomerID="& iCustomerID &" and bDeleted="&getSQLBoolean(false)&" and (iListPageID is null or iListPageID=0) and (bLossePagina="&getSQLBoolean(false)&" or bLossePagina is null) and (iParentID=0 or " & sqlNull("iParentID",iParentID) &")"
if convertBool(bIntranet) then
sql=sql& " and bIntranet="&getSQLBoolean(true)&" "
else
sql=sql& " and (bIntranet="&getSQLBoolean(false)&" or bIntranet is null) "
end if
set rs=db.execute(sql)
getRang=clng(rs(0))
'Response.Write getRang & "<br />" & sql
'Response.End 
set rs=nothing
end function
public function removeRang(byref page)
if convertGetal(page.iListPageID)=0 and (page.bLossePagina=false or isNull(page.bLossePagina)) then
dim sql
sql="update tblPage set iRang=iRang-1 where "& sqlCustId &" and " & sqlNull("iParentID",iId) & " and iRang>"& page.iRang
if page.bIntranet then
sql=sql& " and bIntranet="&getSQLBoolean(true)&" "
else
sql=sql& " and (bIntranet="&getSQLBoolean(false)&" or bIntranet is null) "
end if
db.execute(sql)
clearMenuCache
end if
end function
public property get getLink(includePijltje)
if bContainerPage and isLeeg(sPw) then
getLink="<a href=""#"" style=""cursor: default;"" id='QS_VMENU" & encrypt(iId) & "' onclick=""javascript:return false;"">" & sTitleForMenu & "</a>"
'response.write sTitleForMenu & "<br />"
else
getLink=getClickLink(false)
end if
if includePijltje and not bam then
if subpages(true).count>0 then
getLink=replace(getLink,"<a ","<a class='subcontainer' ",1,-1,1)
end if
end if
end property
public property get getSimpleLink
if customer.bUserFriendlyURL and not isLeeg(sUserFriendlyURL) then
getSimpleLink=sUserFriendlyURL
else
getSimpleLink="default.asp?iId=" & encrypt(iId)
end if
end property
public property get getClickLink(BO)
if BO then
if not isLeeg(sExternalURL) then
getClickLink="<a href='bs_editExternalURL.asp?iId="& encrypt(iId)&"'"
elseif bContainerPage then
getClickLink="<a href='bs_editContainer.asp?iId="& encrypt(iId)&"'"
elseif not isLeeg(sOrderBY) then 'list page!
getClickLink="<a href='bs_listPage.asp?iId="& encrypt(iId)	&"'"
elseif convertGetal(iListPageID)<>0 then
getClickLink="<a href='bs_editListItem.asp?iId="& encrypt(iId)&"'"
else 
getClickLink="<a href='bs_editItem.asp?iId="& encrypt(iId)&"'"
end if
else
if not isLeeg(sExternalURL) then
getClickLink="<a "
if bOpenInNewWindow then getClickLink=getClickLink&" target='_blank'"
getClickLink=getClickLink&" href=""" & replace(sExternalURLPrefix & sExternalURL,"&","&amp;") & """"
if not isLeeg(sRel) then
getClickLink=getClickLink&" rel=""" & sRel & """"
end if
if not isLeeg(sClassname) then
getClickLink=getClickLink&" class=""" & sClassname & """"
end if
elseif customer.bUserFriendlyURL and not isLeeg(sUserFriendlyURL) then
getClickLink="<a href='"&sUserFriendlyURL&"'"
else
getClickLink="<a href='default.asp?iId="& encrypt(iId)&"'"
end if
end if
if statusString=l("offline") then
getClickLink=getClickLink&" style=""color:" & MYQS_offlineLinkColor & """><i>Offline</i>: "
else
getClickLink=getClickLink&" id='QS_VMENU"&encrypt(iId)&"'>"
end if
if convertGetal(iListPageID)<>0 then
getClickLink=getClickLink & quotRep(sDateAndTitle) 
else
getClickLink=getClickLink & sTitleForMenu 
end if
getClickLink=getClickLink & "</a>"
end property
public property get getParentLink
if bContainerPage and isLeeg(sPw) then
getParentLink=sTitle
else
getParentLink=getClickLink(false)
end if
end property
public function copy
if not isLeeg(iId) then
dim copylistitems, cItem
set copylistitems=listitems(false)
iId=null
bHomepage=false
sUserFriendlyURL=""
sCode=""
sTitle =l("copyof") & " " & sTitle
save()
for each cItem in copylistitems
copylistitems(cItem).iListPageID=iId
copylistitems(cItem).copy()
next
set copylistitems=nothing
clearMenuCache
end if
end function
public function resetAllSubPasswords (pageId)
dim page
set page=new cls_page
page.pick(pageId)
if isLeeg(page.iId) then exit function
dim copySubPages, subPage
set copySubPages=page.subPages(false)
for each subPage in copySubPages
copySubPages(subPage).sPw=sPw
copySubPages(subPage).save()
resetAllSubPasswords(copySubPages(subPage).iId)
next
set copySubPages=nothing
clearMenuCache
end function
public function removeAllSubPasswords (pageId)
dim page
set page=new cls_page
page.pick(pageId)
if isLeeg(page.iId) then exit function
dim copySubPages, subPage
set copySubPages=page.subPages(false)
for each subPage in copySubPages
copySubPages(subPage).sPw=""
copySubPages(subPage).save()
removeAllSubPasswords(copySubPages(subPage).iId)
next
set copySubPages=nothing
clearMenuCache
end function
public function moveToFirstSubItem()
dim rs
set rs=db.execute("select iId,sUserFriendlyURL from tblPage where "& sqlCustId &" and " & sqlNull("iParentID",iId) & " and bDeleted="&getSQLBoolean(false)&" and (sPw='"& sPw &"' or sPw is null) order by iRang asc")
if not rs.eof then
dim checkUFL
checkUFL=rs("sUserFriendlyURL")
if not isLeeg(checkUFL) and customer.bUserFriendlyURL then
Response.Redirect (checkUFL)
else
Response.Redirect ("default.asp?iId="&encrypt(rs(0)))
end if
else
Response.Redirect ("default.asp")
end if
end function
public function fastlistitems
set fastlistitems=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
if not isleeg(sOrderBY) then
sql="select * from tblPage where bDeleted="&getSQLBoolean(false)&" and iListPageID="& iID
sql=sql&" and (dOnlineFrom is null or dOnlineFrom<="&getSQLDateFunction&") and (dOnlineUntill is null or dOnlineUntill>="&getSQLDateFunction&")"
sql=sql&" order by "&sOrderBY
if instr(sOrderBY,"updated")=0 then
sql=sql& ", updatedTS DESC"
end if
dim rs,sql, page
set rs=db.execute(sql)
while not rs.eof
set page=new cls_fastpage
page.iId=rs("iId")
page.bHideDate=bHideDate
page.sLPExternalURL=rs("sLPExternalURL")
page.sTitle=rs("sTitle")
page.bLPExternalOINW=rs("bLPExternalOINW")
page.sValue=rs("sValue")
page.iFeedId=rs("iFeedId")
page.sLPExternalURL=rs("sLPExternalURL")
page.dPage=rs("dPage")
page.updatedTS=rs("updatedTS")
page.sItemPicture=rs("sItemPicture")
page.sUserfriendlyURL=rs("sUserfriendlyURL")
page.sLPIC=rs("sLPIC")
page.iPMlocation=rs("iPMlocation")
'page.sUrlRRSImage=rs("sUrlRRSImage")
fastlistitems.Add page.iId, page
set page=nothing
rs.movenext
wend
end if
end if
end function
public function listitems(onlineOnly)
set listitems=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
if not isleeg(sOrderBY) then
sql="select iID from tblPage where bDeleted="&getSQLBoolean(false)&" and iListPageID="& iID
if onlineOnly then
sql=sql&" and (dOnlineFrom is null or dOnlineFrom<="&getSQLDateFunction&") and (dOnlineUntill is null or dOnlineUntill>="&getSQLDateFunction&")"
end if
sql=sql&" order by "&sOrderBY
if instr(sOrderBY,"updated")=0 then
sql=sql& ", updatedTS DESC"
end if
dim rs,sql, page
set rs=db.execute(sql)
while not rs.eof
set page=new cls_page
page.pick(rs(0))
listitems.Add page.iId, page
set page=nothing
rs.movenext
wend
set rs=nothing
end if
end if
end function
Public property get	statusString
if not isLeeg(iListPageID) then
if isBetween(dOnlineFrom,date,dOnlineUntill) then
statusString=l("online")
else
statusString=l("offline")
end if
elseif bOnline then
statusString=l("online")
else
statusString=l("offline")
end if
end property
Public Property get sDateAndTitle
if isLeeg(dPage) then
sDateAndTitle=sTitle
else
if convertBool(bHideDate) then
sDateAndTitle=sTitle
else
sDateAndTitle=convertEuroDate(dPage)& ": " & sTitle
end if
end if
end property
Public sub addHit()
if not isLeeg(iId) then
dim rs
set rs=db.execute("update tblPage set iHits="& convertGetal(iHits)+1 &" where iID="& iID)
set rs=nothing
end if
end sub
Public sub addHitRSS()
if not isLeeg(iId) then
dim rs
set rs=db.execute("update tblPage set iHitsRSS="& convertGetal(iHitsRSS)+1 &" where iID="& iID)
set rs=nothing
end if
end sub
Public sub addVisit()
if not isLeeg(iId) then
if isLeeg(session("visitorLoaded"&iID)) then
dim rs
set rs=db.execute("update tblPage set iVisitors="& convertGetal(iVisitors)+1 &" where iID="& iID)
set rs=nothing
session("visitorLoaded"&iID)=true
end if
end if
end sub
Public function catalog
set catalog=new cls_catalog
catalog.pick(iCatalogId)
end function 
Public function Feed
set Feed=new cls_Feed
Feed.pick(iFeedId)
end function 
Public function form
set form=new cls_form
form.pick(iFormID)
end function 
Public function deleteListItems
dim clistitems, lItem
set clistitems=listitems(false)
for each lItem in clistitems
clistitems(lItem).bDeleted=true
clistitems(lItem).save()
next
set clistitems=nothing
clearMenuCache
end function
public function template
set template=new cls_template
template.pick(iTemplateID)
end function
Public function buildTemplate
on error resume next
dim cTemplate
set cTemplate=template
if isLeeg(iId) then
	set cTemplate=customer.defaulttemplateObj
else
	if isLeeg(cTemplate.iId) then set cTemplate=customer.defaulttemplateObj
end if
if isLeeg(iTemplateID) then
if convertStr(session("iTemplateID"))="0" then
exit function
end if
if convertStr(Request.Cookies("iTemplateID"))="0" then
exit function
end if
if not isLeeg(session("iTemplateID")) then
cTemplate.pick(session("iTemplateID"))
end if
if convertGetal(customer.i404TemplateID)<>0 then
cTemplate.pick(customer.i404TemplateID)
end if
if not isLeeg(Request.Cookies("iTemplateID")) then
cTemplate.pick(Request.Cookies("iTemplateID"))
end if
end if
if not isLeeg(logon.currentPW) then
if isNumeriek(decrypt(Request.QueryString("previewTemplate"))) then
session("iTemplateID")=decrypt(Request.QueryString("previewTemplate"))
cTemplate.pick(session("iTemplateID"))
elseif Request.QueryString("previewTemplate")="false" then
session("iTemplateID")=""
redirectToHP()
end if
end if
if not isLeeg(cTemplate.iId) then
dim sPage
if pagetoemail then
sPage=treatConstants(cTemplate.sEmailValue,true)
elseif printReplies then
sPage=treatConstants(cTemplate.sPrintValue,true)
elseif convertBool(cTemplate.bCompress) then
sPage=treatConstants(cTemplate.sCompressValue,true)
else
sPage=treatConstants(cTemplate.sValue,true)
end if
sPage=Replace(sPage, "[BANNER]", convertStr(customer.sBannerMenu),1,-1,1)
sPage=Replace(sPage, "[QSHIGHLIGHTS]", convertStr(customer.sRightBanner),1,-1,1)
sPage=Replace(sPage, "[QSHIGHLIGHTSLABEL]", convertStr(customer.sHighlights),1,-1,1)
sPage=Replace(sPage, "[QSCONTACTINFOLABEL]", convertStr(customer.sContactInfo),1,-1,1)
sPage=Replace(sPage, "[QSSEARCHLABEL]", l("search"),1,-1,1)
sPage=Replace(sPage, "[QSSITEFOOTER]", convertStr(customer.sFooter),1,-1,1)
dim aantalDagen
aantalDagen=customer.aantalDagen
'page specific
sPage=Replace(sPage, "[C_VIRT_DIR]", C_VIRT_DIR,1,-1,1)
sPage=Replace(sPage, "[C_DIRECTORY_QUICKERSITE]", C_DIRECTORY_QUICKERSITE,1,-1,1)
sPage=Replace(sPage, "[HEADER]", convertStr(customer.sTopheader),1,-1,1)
sPage=Replace(sPage, "[PAGEBODY]", insertMedia(pageBody),1,-1,1)
sPage=Replace(sPage, "[PAGEHEADER]", customer.sHeader & vbcrlf & sHeader,1,-1,1)
sPage=Replace(sPage, "[PAGETITLE]", pageTitle,1,-1,1)
dim bfVMENU
if instr(sPage,"[QS_ARTISTEER_FULLVMENU_V24]")<>0	then
bfVMENU=replace(qs_artisteer_full_menu(null),"<ul class='artmenu'><li><","<ul class='art-vmenu'><li><",1,-1,1)
bfVMENU=treatCI(bfVMENU)
sPage=Replace(sPage, "[QS_ARTISTEER_FULLVMENU_V24]", bfVMENU,1,-1,1)
sPage=replace(sPage,"QS_VMENU","",1,-1,1)
end if
if instr(sPage,"[QS_ARTISTEER_FULLVMENU_V3]")<>0	then 
bfVMENU=replace(qs_artisteer_full_menu(null),"<ul class='artmenu'><li><","<ul class='art-vmenu'><li><",1,-1,1)
bfVMENU=art3XVMENUFIX(bfVMENU)
sPage=Replace(sPage, "[QS_ARTISTEER_FULLVMENU_V3]", bfVMENU,1,-1,1)
sPage=replace(sPage,"QS_VMENU","",1,-1,1)
end if
if instr(sPage,"[QS_ARTISTEER_FULLVMENU_V31]")<>0	then 
bfVMENU=replace(qs_artisteer_full_menu(null),"<ul class='artmenu'><li><","<ul class='art-vmenu'><li><",1,-1,1)
bfVMENU=art3XVMENUFIX(bfVMENU)
bfVMENU=art31MENUFIX(bfVMENU)
sPage=Replace(sPage, "[QS_ARTISTEER_FULLVMENU_V31]", bfVMENU,1,-1,1)
sPage=replace(sPage,"QS_VMENU","",1,-1,1)
end if
if instr(sPage,"[PAGEMENU]")<>0	then sPage=Replace(sPage, "[PAGEMENU]", showMenu,1,-1,1)
if instr(sPage,"[JS_HMENU1]")<>0	then sPage=Replace(sPage, "[JS_HMENU1]", JStemplateHMENU(1),1,-1,1)
if instr(sPage,"[JS_HMENU2]")<>0	then sPage=Replace(sPage, "[JS_HMENU2]", JStemplateHMENU(2),1,-1,1)
if instr(sPage,"[JS_VMENU]")<>0	then sPage=Replace(sPage, "[JS_VMENU]", JStemplateVMENU,1,-1,1)
if instr(sPage,"[QS_BOOTSTRAPMENU_3]")<>0	then sPage=Replace(sPage, "[QS_BOOTSTRAPMENU_3]", getBootstrapMenu(3),1,-1,1)
if instr(sPage,"[QS_MAINMENU]")<>0	then sPage=Replace(sPage, "[QS_MAINMENU]", qs_mainmenu,1,-1,1)
if instr(sPage,"[QS_INTRANETMENU]")<>0	then sPage=Replace(sPage, "[QS_INTRANETMENU]", qs_intranetmenu,1,-1,1)
if instr(sPage,"[QS_ARTISTEER_INTRANETMENU]")<>0	then sPage=Replace(sPage, "[QS_ARTISTEER_INTRANETMENU]", qs_artisteer_menu(true),1,-1,1)
if instr(sPage,"[QS_ARTISTEER_MENU]")<>0	then sPage=Replace(sPage, "[QS_ARTISTEER_MENU]", qs_artisteer_menu(false),1,-1,1)
if instr(sPage,"[QS_ARTISTEER_FULLMENU]")<>0	then sPage=Replace(sPage, "[QS_ARTISTEER_FULLMENU]", qs_artisteer_full_menu(null),1,-1,1)
if instr(sPage,"[QS_ARTISTEER_FULLMENU_V22]")<>0	then 
sPage=Replace(sPage, "[QS_ARTISTEER_FULLMENU_V22]", replace(qs_artisteer_full_menu(null),"<ul class='artmenu'><li><","<ul class='art-menu'><li><",1,-1,1),1,-1,1)
end if
if instr(sPage,"[QS_ARTISTEER_FULLMENU_V3]")<>0	then 
sPage=Replace(sPage, "[QS_ARTISTEER_FULLMENU_V3]", replace(qs_artisteer_full_menu(null),"<ul class='artmenu'><li><","<ul class='art-hmenu'><li><",1,-1,1),1,-1,1)
end if
if instr(sPage,"[QS_ARTISTEER_FULLMENU_V31]")<>0	then 
sPage=Replace(sPage, "[QS_ARTISTEER_FULLMENU_V31]", replace(art31MENUFIX(qs_artisteer_full_menu(null)),"<ul class='artmenu'><li><","<ul class='art-hmenu'><li><",1,-1,1),1,-1,1)
end if
dim helpArtist,helpArtist2
if instr(sPage,"[QS_ARTISTEER_FULLMENU_V4]")<>0	then 
helpArtist=qs_artisteer_parentlink
qs_artisteer_parentlink="{TITLE}"
helpArtist2=replace(art31MENUFIX(qs_artisteer_full_menu(null)),"<ul class='artmenu'><li><","<ul class='art-hmenu'><li><",1,-1,1)
helpArtist2=replace(helpArtist2,"</span>","",1,-1,1)
helpArtist2=replace(helpArtist2,"<span class=""t"">","",1,-1,1)
sPage=Replace(sPage, "[QS_ARTISTEER_FULLMENU_V4]", helpArtist2,1,-1,1)
qs_artisteer_parentlink=helpArtist
end if
if instr(sPage,"[QS_ARTISTEER_FULLVMENU_V4]")<>0	then 
helpArtist=qs_artisteer_parentlink
qs_artisteer_parentlink="{TITLE}"
bfVMENU=replace(qs_artisteer_full_menu(null),"<ul class='artmenu'><li><","<ul class='art-vmenu'><li><",1,-1,1)
bfVMENU=art3XVMENUFIX(bfVMENU)
bfVMENU=art31MENUFIX(bfVMENU)
bfVMENU=replace(bfVMENU,"</span>","",1,-1,1)
bfVMENU=replace(bfVMENU,"<span class=""t"">","",1,-1,1)
sPage=Replace(sPage, "[QS_ARTISTEER_FULLVMENU_V4]", bfVMENU,1,-1,1)
sPage=replace(sPage,"QS_VMENU","",1,-1,1)
qs_artisteer_parentlink=helpArtist
end if
if instr(sPage,"[PAGEBREADCRUMBS]")<>0	then sPage=Replace(sPage, "[PAGEBREADCRUMBS]", showBreadCrumbs,1,-1,1)
if instr(sPage,"[CANONICALLINK]")<>0	then sPage=Replace(sPage, "[CANONICALLINK]", "",1,-1,1)
sPage=Replace(sPage, "[PAGEFEED]", pageFeed,1,-1,1)
sPage=Replace(sPage, "[PAGECATALOG]", pageCatalog,1,-1,1)
sPage=Replace(sPage, "[PAGELIST]", pageListPage,1,-1,1)
sPage=Replace(sPage, "[PAGEFORM]", pageForm,1,-1,1)
sPage=Replace(sPage, "[PAGECREATED_TS]", createdTS,1,-1,1)
sPage=Replace(sPage, "[PAGEUPDATED_TS]", updatedTS,1,-1,1)
sPage=Replace(sPage, "[PAGENMBRVISITS]", iVisitors,1,-1,1)
sPage=Replace(sPage, "[PAGENMBRHITS]", iHits,1,-1,1)
sPage=Replace(sPage, "[PAGENMBRVISITS_DAY]", round(convertGetal(iVisitors/aantalDagen)),1,-1,1)
sPage=Replace(sPage, "[PAGENMBRHITS_DAY]", round(convertGetal(iHits/aantalDagen)),1,-1,1)
sPage=Replace(sPage, "[PAGEhref]", "default.asp?iId="&EnCrypt(iId),1,-1,1)
sPage=Replace(sPage, "[PAGEhref_EMAIL]", C_DIRECTORY_QUICKERSITE & "/mailpage.asp?iId="&EnCrypt(iId),1,-1,1)
sPage=Replace(sPage, "[PAGEhref_PRINT]", C_DIRECTORY_QUICKERSITE & "/printpage.asp?iId="&EnCrypt(iId),1,-1,1)
sPage=Replace(sPage, "[TITLETAG]", removeHTML(showTitle),1,-1,1)
sPage=Replace(sPage, "[RSSLINK]", linkRSS,1,-1,1)
sPage=Replace(sPage, "[RSSLINKART]", sArtRSSLink,1,-1,1)
sPage=Replace(sPage, "[METATAG_KEYWORDS]", quotRep(convertStr(showKeywords)),1,-1,1)
sPage=Replace(sPage, "[METATAG_DESCRIPTION]", quotRep(convertStr(showDescription)),1,-1,1)
sPage=Replace(sPage, "[METATAG_REFRESH]",metatag_refresh,1,-1,1)
if customer.hasFavicon then
sPage=Replace(sPage, "[FAVICON_LINK]","<link rel=""shortcut icon"" type=""image/x-icon"" href=""" & C_VIRT_DIR &  Application("QS_CMS_userfiles") & "favicon.ico"" />",1,-1,1)	 
else
sPage=Replace(sPage, "[FAVICON_LINK]","",1,-1,1)
end if
'fixed & required parts LEAVE THIS
sPage=Replace(sPage, "[METATAG_GENERATOR]", MYQS_GENERATOR,1,-1,1)
sPage=Replace(sPage, "[METATAG_CHARSET]", QS_CHARSET,1,-1,1)
'metatags
sPage=Replace(sPage, "[METATAG_AUTHOR]", convertStr(quotRep(customer.webmaster)),1,-1,1)
sPage=Replace(sPage, "[METATAG_COPYRIGHT]", convertStr(quotRep(customer.copyRight)),1,-1,1)
'elements for search form
sPage=Replace(sPage, "[SF_HIDDENFIELD]","<input type='hidden' name='pageAction' value='search' />",1,-1,1)
sPage=Replace(sPage, "[SF_ACTIONURL]","default.asp",1,-1,1)
sPage=Replace(sPage, "[SF_TEXTVALUE]", convertStr(quotrep(search.returnValue)),1,-1,1)
'extra
sPage=Replace(sPage, "[SITENAME]", convertStr(customer.siteName),1,-1,1)
sPage=Replace(sPage, "[SITESLOGAN]", convertStr(customer.sSiteSlogan),1,-1,1)
sPage=Replace(sPage, "[SITEMAP]", showSiteMap,1,-1,1)
sPage=Replace(sPage, "[SITEURL]", convertStr(customer.sUrl),1,-1,1)
sPage=Replace(sPage, "[WEBMASTEREMAIL]", convertStr(customer.webmasterEmail),1,-1,1)
sPage=Replace(sPage, "[GOOGLEANALYTICS]", convertStr(customer.googleAnalytics),1,-1,1)
sPage=Replace(sPage, "[href_SITEMAP]", "default.asp?pageAction=sitemap",1,-1,1)
sPage=Replace(sPage, "[href_HOME]", "default.asp?iId="&EnCrypt(getHomepage.iId),1,-1,1)

sPage=Replace(sPage, "</body>",customer.getArrowJS & "</body>",1,-1,1)



'insert blocks
sPage=replaceBlocks(sPage)
'insert constants
sPage=replace(treatConstants(sPage,true)," & "," &amp; ",1,-1,1) 
'include theme
sPage=Replace(sPage, "[PAGETHEME]", pageTheme,1,-1,1)
'insert system messages
if message.hasMessages then
sPage=Replace(sPage, "</body>",message.showAlert & "</body>",1,-1,1)
end if
if customer.bPeelEnabled then
sPage=replace(sPage,"<body>","<body>" & customer.getPeelDiv,1,-1,1)
end if
if customer.bCookieWarning then
sPage=Replace(sPage, "</body>",customer.sCookieJS & "</body>",1,-1,1)
end if
'insert custom colorbox
if instr(sPage,"QSCCB_")<>0 then
dim startpos,endpos,checkCCB,cccCode,arrCCC
set checkCCB=server.createObject("scripting.dictionary")
startpos=instr(sPage,"QSCCB_")
while startpos<>0 
endpos=instr(startpos,sPage," ")
cccCode=mid(sPage,startpos,endpos-startpos)
cccCode=replace(cccCode,"'","",1,-1,1)
cccCode=replace(cccCode,"""","",1,-1,1)
cccCode=replace(cccCode," ","",1,-1,1)
if not checkCCB.exists(cccCode) and len(cccCode)<17 then
on error resume next
err.clear()
arrCCC=split(cccCode,"_")
sCBExtUrl=sCBExtUrl&"<script type=""text/javascript"">" & vbcrlf
sCBExtUrl=sCBExtUrl&"$(document).ready(function(){" & vbcrlf
sCBExtUrl=sCBExtUrl&"$(""." & cccCode & """).colorbox({close: """ & quotrep(l("close")) & """, width:""" & arrCCC(1) & """, height:""" & arrCCC(2) & """, iframe:true}); " & vbcrlf
sCBExtUrl=sCBExtUrl&"});" & vbcrlf & "</script>" & vbcrlf
checkCCB.Add cccCode,""
if err.number<>0 then sCBExtUrl=""
on error goto 0
end if
startpos=instr(endpos,sPage,"QSCCB_")
wend
set checkCCB=nothing
end if
sPage=Replace(sPage, "[JAVASCRIPT]",dumpJavaScript(true),1,-1,1)
sPage=Replace(sPage, "[JS_JAVASCRIPT]",dumpJavaScript(false),1,-1,1)
'include pagerender time
sPage=Replace(sPage, "[PAGERENDERTIME]", PrintTimer(startTimer),1,-1,1)
'execute custom script?
if runAPP then
if not isLeeg(sApplication) and customer.bApplication and not pagetoemail then
if instr(1,sPage,"[PAGEAPPLICATION]",vbTextCompare)<>0 then
dim twoParts
twoParts=split(sPage,"[PAGEAPPLICATION]")
Response.write twoParts(0)
'Response.Flush 
'execute
server.execute(treatConstants(sApplication,true))
Response.write twoParts(1)
'cleanup
set cTemplate=nothing
'stop execute further lines!
cleanUPASP
'execute afterPageLoadScript
run(execAfterPageLoad)
Response.End 
end if
end if
end if
sPage=Replace(sPage, "[PAGEAPPLICATION]", "",1,-1,1)
sPage=replace(sPage,"|R|R|R|","[",1,-1,1)
if pagetoemail then
buildTemplate=sPage
exit function
else
Response.Write sPage
'Response.Flush 
'cleanup
set cTemplate=nothing
logReferer()
'caching
'response.write "bLoadInCache: " & bLoadInCache
'response.write customer.sQSUrl & "/default.asp?iId=" & encrypt(iId)
if bLoadInCache and convertBool(application("doLoadFromCache")) then
dim oXMLHTTP
Set oXMLHTTP = server.createobject("MSXML2.ServerXMLHTTP")
oXMLHTTP.open "GET", customer.sQSUrl & "/default.asp?iId=" & encrypt(iId)
oXMLHTTP.send
sPageCache=oXMLHTTP.responseText
if not isLeeg(sPageCache) then
if len(sPageCache)<60000 then
db.execute("update tblPage set sPageCache='" & replace(sPageCache,"'","''",1,-1,1) & "' where iId="& iId)
end if
end if
set oXMLHTTP=nothing
end if
application("doLoadFromCache")=true
'stop execute further lines!
cleanUPASP
'execute afterPageLoadScript
run(execAfterPageLoad)
'if err.number<>0 then
'	response.write "<script type=""text/javascript"">" & quotrepJS(err.description) & "</script>"
'end if
Response.End 
end if
end if
on error goto 0
end function
function treatCI(v)
'response.clear 
'response.write server.htmlencode(v)
'response.write "<br />"
dim sPOS
v=replace(v,"onclick=""javascript:return false;""","",1,-1,1)
v=replace(v,"style=""cursor: default;""","",1,-1,1)
v=replace(v,"class='subcontainer'","",1,-1,1)
v=replace(v,"""#""","""#"" onclick=""javascript:location.assign('default.asp?iId='+this.id);return false;""",1,-1,1)
v=replace(v,"'#'","'#' onclick=""javascript:location.assign('default.asp?iId='+this.id);""",1,-1,1)
'v=replace(v," id='"," id='QVMENU",1,-1,1)
'v=replace(v,"""#""","""#"" onclick=""javascript:alert(this.id);""",1,-1,1)
treatCI=v
'response.write server.htmlencode(v)
'response.end 
end function
public property get showTitle
if not isLeeg(sSEOtitle) then
showTitle=sSEOtitle
elseif not isLeeg(sTitle) then
showTitle=sTitle& " | "  & customer.siteTitle
else
showTitle=customer.siteTitle
end if
showtitle=treatConstants(showtitle,true)
end property
public property get showKeywords
if not isLeeg(sKeywords) then
showKeywords=sKeywords
else
showKeywords=customer.keywords
end if
showKeywords=treatConstants(showKeywords,true)
end property
public property get showDescription
if not isLeeg(sDescription) then
showDescription=sDescription
else
showDescription=customer.sDescription
end if
showDescription=treatConstants(showDescription,true)
end property
public function makeNormalPage
dim copyListitems,lItem
set copyListitems=listitems(true)
for each lItem in copyListitems
sValue=sValue& "<b>" & copyListitems(lItem).sDateAndTitle &"</b>"
sValue=sValue&copyListitems(lItem).sValue
copyListitems(lItem).bDeleted=true
copyListitems(lItem).save()
next
set copyListitems=nothing
sOrderby=""
makeNormalPage=Save
clearMenuCache
end function
public function linkRSS()
dim crss
crss=rssLink
if not isLeeg(crss) then
linkRSS="<link rel=""alternate"" type=""application/rss+xml"" title=""RSS"" href=" & """" & treatConstants(crss,true) & """" &" />"
end if
end function
public function rssLink
rssLink=customer.sDefaultRSSLink
if isNumeriek(iId) and bOnline and not bDeleted then
'check password
if logon.logonItem(Request.Cookies(encrypt(iId)),me) then
if bIntranet then
if not logon.authenticatedIntranet or not convertBool(customer.intranetUse) then
exit function
end if
end if
dim showLink
showLink=false
'listpage?
if not isLeeg(sOrderBY) and bPushRSS then
showLink=true
end if
'themepage?
if convertGetal(iThemeID)<>0 then
dim copytheme
set copytheme=theme
if copytheme.bPushRSS and copytheme.bOnline then
showLink=true
end if
end if
'catalog?
if convertGetal(iCatalogId)<>0 then
dim copyCatalog
set copyCatalog=catalog
if copyCatalog.bPushRSS and copyCatalog.bOnline then
showLink=true
end if
end if
if showLink then 
rssLink=C_DIRECTORY_QUICKERSITE & "/rss.asp?iId=" & encrypt(iId)
end if
if not isLeeg(sRSSLink) then
rssLink=sRSSLink
end if
end if
end if
end function
public property get metatag_refresh
if convertGetal(iReload)>0 then
if bIntranet and not logon.authenticatedIntranet then exit property
metatag_refresh="<META HTTP-EQUIV="& """" & "refresh"& """" & " CONTENT=" & """" & iReload &";URL=" 
if not isLeeg(sRedirectTo) then
metatag_refresh=metatag_refresh&sRedirectTo
else
metatag_refresh=metatag_refresh&"default.asp?iID="&EnCrypt(iID)
end if
metatag_refresh=metatag_refresh& """" & " />"
end if
end property
public function theme
set theme=new cls_theme
theme.pick(iThemeID)
end function
Public property get bCanBeConvertedToFP
bCanBeConvertedToFP=isNumeriek(iID) and secondAdmin.bPageBody and subPages(false).count=0 and not bLossePagina and convertGetal(iListPageID)=0 and isLeeg(sExternalURLPrefix)
end property
Public function convertToFP
convertToFP=false
if bCanBeConvertedToFP then
dim sql
sql="update tblPage set iRang=iRang-1 where "& sqlCustId &" and " & sqlNull("iParentID",iParentID) & " and iRang>"& iRang
if bIntranet then
sql=sql& " and bIntranet="&getSQLBoolean(true)&" "
else
sql=sql& " and (bIntranet="&getSQLBoolean(false)&" or bIntranet is null) "
end if
'Response.Write sql
'Response.End 
db.execute(sql)
clearMenuCache
iRang=0
iParentid=null
bLossePagina=true
convertToFP=save()
end if
end function
public function copyToCustomerID(oCID)
if isNumeriek(iId) then
iId	= null
overruleCID	= oCID
iCustomerID = overruleCID
save()
end if
end function
public property get sGetTitle
if isLeeg(sPageTitle) then
sGetTitle=sTitle
else
sGetTitle=sPageTitle
end if
end property
public property get sArtRSSLink
if not isLeeg(rssLink) then
sArtRSSLink="<a target=""_blank"" href=""" & rssLink & """ class=""art-rss-tag-icon"" title=""RSS""></a>"
end if
end property
public function updater
set updater=new cls_contact
if convertGetal(iUpdatedBy)<>0 then
updater.pick(iUpdatedBy)
end if
end function
public function replaceBlocks(sValue)
replaceBlocks=sValue
if instr(replaceBlocks,"[PAGE_BLOCK0")=0 then exit function
if not isLeeg(sProp01) then
replaceBlocks=replace(replaceBlocks,"[PAGE_BLOCK01]",sProp01,1,-1,1)
elseif not isLeeg(customer.sProp01) then
replaceBlocks=replace(replaceBlocks,"[PAGE_BLOCK01]",customer.sProp01,1,-1,1)
end if
if not isLeeg(sProp02) then
replaceBlocks=replace(replaceBlocks,"[PAGE_BLOCK02]",sProp02,1,-1,1)
elseif not isLeeg(customer.sProp02) then
replaceBlocks=replace(replaceBlocks,"[PAGE_BLOCK02]",customer.sProp02,1,-1,1)
end if
if not isLeeg(sProp03) then
replaceBlocks=replace(replaceBlocks,"[PAGE_BLOCK03]",sProp03,1,-1,1)
elseif not isLeeg(customer.sProp03) then
replaceBlocks=replace(replaceBlocks,"[PAGE_BLOCK03]",customer.sProp03,1,-1,1)
end if
if not isLeeg(sProp04) then
replaceBlocks=replace(replaceBlocks,"[PAGE_BLOCK04]",sProp04,1,-1,1)
elseif not isLeeg(customer.sProp04) then
replaceBlocks=replace(replaceBlocks,"[PAGE_BLOCK04]",customer.sProp04,1,-1,1)
end if
if not isLeeg(sProp05) then
replaceBlocks=replace(replaceBlocks,"[PAGE_BLOCK05]",sProp05,1,-1,1)
elseif not isLeeg(customer.sProp05) then
replaceBlocks=replace(replaceBlocks,"[PAGE_BLOCK05]",customer.sProp05,1,-1,1)
end if
if not isLeeg(sProp06) then
replaceBlocks=replace(replaceBlocks,"[PAGE_BLOCK06]",sProp06,1,-1,1)
elseif not isLeeg(customer.sProp06) then
replaceBlocks=replace(replaceBlocks,"[PAGE_BLOCK06]",customer.sProp06,1,-1,1)
end if
if not isLeeg(sProp07) then
replaceBlocks=replace(replaceBlocks,"[PAGE_BLOCK07]",sProp07,1,-1,1)
elseif not isLeeg(customer.sProp07) then
replaceBlocks=replace(replaceBlocks,"[PAGE_BLOCK07]",customer.sProp07,1,-1,1)
end if
if not isLeeg(sProp08) then
replaceBlocks=replace(replaceBlocks,"[PAGE_BLOCK08]",sProp08,1,-1,1)
elseif not isLeeg(customer.sProp08) then
replaceBlocks=replace(replaceBlocks,"[PAGE_BLOCK08]",customer.sProp08,1,-1,1)
end if
dim iB
for iB=1 to 8
replaceBlocks=replace(replaceBlocks,"[PAGE_BLOCK0" & iB & "]","",1,-1,1)
next
end function
public sub showFromCache()
'algemeen: een pagina met paswoord mag nooit gecached worden
'de caching zelf dient te gebueren door een http request
if not convertBool(customer.bUseCachingForPages) then exit sub
if convertbool(bNocache) then exit sub
'alleen pagina's met userfriendly url
if isLeeg(sUserFriendlyURL) then exit sub
'paginas met custom scripting niet cachen
if not isLeeg(sApplication) then exit sub
'nooit pagina's met formulieren cachen
if convertGetal(iFormID)<>0 then exit sub
'someone is logged in to the intranet
if not isLeeg(Request.Cookies("sMode")) then exit sub
if convertGetal(logon.contact.iId) then exit sub
'nooit pagina's cachen achter paswoord
if not isLeeg(sPw) then exit sub
'no cache available
if isLeeg(sPageCache) then 
bLoadInCache=true
exit sub
end if
'pagina's met een poll nooit cache tonen
if instr(sPageCache,"getVote(this")<>0 then 
exit sub
end if
'pagina's met een popup nooit cachen
if instr(sPageCache,"href:sGetUrl")<>0 then 
exit sub
end if
'pagina's met een newsletter nooit cachen
if instr(sPageCache,"getSub(")<>0 then
exit sub
end if
'counters bijhouden
AddHit()
addVisit()
'response.write "<div style=""background-color:Red;color:White;padding:10px"">" & PrintTimer(startTimer) & "</div>"
response.write sPageCache & vbcrlf &"<!-- loaded from cache in " & PrintTimer(startTimer) & " -->"
logReferer()
'stop execute further lines!
cleanUPASP
'execute afterPageLoadScript
run(execAfterPageLoad)
response.end
end sub
public sub clearPageCache
on error resume next
if customer.bUseCachingForPages then
dim rs
set rs=db.execute("update tblPage set sPageCache=null where iId=" & convertGetal(iId))
set rs=nothing
application("doLoadFromCache")=false
end if
on error goto 0
end sub
public function getAccordeonList
if bAccordeon then
dim h,fli,li,liCounter,isLink
liCounter=0
set fli=fastlistitems
h="<div class=""QSAccordion"" id=""accordion" & encrypt(iId) & """>"
for each li in fli
fli(li).bHideDate=selectedPage.bHideDate
if Request.QueryString ("item")=encrypt(li) then
selectedPage.sSEOtitle=quotrep(fli(li).sTitle & " | "  & customer.siteTitle)
end if
isLink=not isLeeg(fli(li).sLPExternalURL) and fli(li).sLPExternalURL<>"http://"
if isLink then
h=h & "<div class=""QSAccordionHeader""><a name=""" & encrypt(li) & """></a>" & fli(li).sDateAndTitle & "</div>"
h=h & "<div class=""QSAccordionContent""><a href=""" & fli(li).sLPExternalURL & """" 
if convertBool(fli(li).bLPExternalOINW) then
h=h & " target=""_blank"""
end if
h=h & ">" & fli(li).sDateAndTitle & "</a></div>"
else
h=h & "<div class=""QSAccordionHeader""><a name=""" & encrypt(li) & """></a>" & fli(li).sDateAndTitle & "</div>"
h=h & "<div class=""QSAccordionContent"">" & fli(li).listitemPicIMGTag & fli(li).page.insertMedia(fli(li).sValue)
if not isLeeg(fli(li).iFeedId) then h=h & fli(li).Feed.build()
h=h & "</div>" 
end if
if convertGetal(decrypt(Request.QueryString ("item")))=convertGetal(li) then iLPDEFAULTOpenerQS=liCounter
liCounter=liCounter+1
next
h=h & "</div>"
set fli=nothing
getAccordeonList=treatconstants(h,true)
if isLeeg(Request.QueryString ("item")) then
if not bOpenOnload then
if convertGetal(iLPOpenByDefault)=0 then 
iLPDEFAULTOpenerQS="""false"""
end if
end if
end if
end if
end function

	public function deleteListItemImage
		on error resume next	
		if not isLeeg(sItemPicture) then
			dim fso
			set fso=server.createobject("scripting.filesystemobject")
			on error resume next
			fso.deletefile server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles")) & "/listitemimages/" & iId & "." & sItemPicture
			on error goto 0
			set fso=nothing	
			
			sItemPicture=""
			save()			
		end if
		on error goto 0
	end function
	
	public function setOnlineListItemsOnline
	
		if convertGetal(iListPageID)<>0 then
			
			if isBetween(dOnlineFrom,date,dOnlineUntill) then				
				bOnline=true
			end if
			
		end if
		
	end function
	
	public property get listitemPicIMGTag
	
		if convertBool(customer.bListItemPic) then

			if not isLeeg(sItemPicture) then
			
				dim cStyle
				select case sLPIC
					case "fp"
						cStyle=" style=""margin:10px 0px 10px 0px;width:100%"" "
					case "al"
						cStyle=" style=""margin:5px 20px 20px 0px;width:48%;float:left"" "
					case "ar"
						cStyle=" style=""margin:5px 0px 20px 20px;width:48%;float:right"" "
					
				end select	
				
				listitemPicIMGTag="<img src=""" & C_VIRT_DIR &  Application("QS_CMS_userfiles") & "listitemimages/" & iId & "." & sItemPicture & """ " & cStyle & " class=""ListItemPictureCSS"" />"
			end if
		
		end if
		
	end property
	
	public function mediafiles
	
		set mediafiles=server.createobject("scripting.dictionary")		
	
		dim fso : set fso=server.createobject("scripting.filesystemobject")
		
		if fso.folderexists(server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & "pagemedia/" & iId)) then
			
			dim file,files : set files=fso.getfolder(server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & "pagemedia/" & iId)).files
			
			for each file in files
				mediafiles.add file.name, file.name
			next
			
			set files=nothing			
			
		end if
	
		set fso=nothing
	
	end function
	
	public sub removeallmedia
	
		dim fso : set fso=server.createobject("scripting.filesystemobject")	
		fso.deletefolder(server.MapPath (C_VIRT_DIR &  Application("QS_CMS_userfiles") & "pagemedia/" & iId))
	
	end sub
	
	public function insertMedia(value)
		
		if convertGetal(iPMlocation)<>0 then

			dim cMediafiles : set cMediafiles=mediafiles
			
			if cMediafiles.count>0 then

				dim medf,medfstr	
				
				dim pad4, padding, perc, tresh, imgSize : imgSize=600 : padding="" : countB=false : pad4="imgPadding"

				select case cMediafiles.count
				
					case 1 : perc="100" : tresh=1 : imgSize=1920 : padding="padding:0px 0px 0px 0px;" : pad4=""
					case 2 : perc="50" 	: tresh=1 : imgSize=1920
					case 3 : perc="33" 	: tresh=1 
					case 4 : perc="50" 	: tresh=1 
					case 5 : perc="33" 	: tresh=1
					case 6 : perc="33" 	: tresh=2
					case 7 : perc="50" 	: tresh=4
					case 8 : perc="25" 	: tresh=2
					case 9 : perc="33" 	: tresh=3
					case 10 : perc="50" : tresh=5
					case 11 : perc="33" : tresh=4
					case 12 : perc="33" : tresh=4
					case 13 : perc="50" : tresh=7
					case 14 : perc="25" : tresh=5
					case 15 : perc="33" : tresh=5
					case 16 : perc="25" : tresh=4
					case 17 : perc="33" : tresh=6
					case 18 : perc="33" : tresh=6
					case 19 : perc="25" : tresh=5
					case 20 : perc="25" : tresh=5
					case 21 : perc="33" : tresh=7
					case 22 : perc="33" : tresh=8
					case 23 : perc="25" : tresh=6
					case 24 : perc="33" : tresh=8
					case 25 : perc="20" : tresh=5
					case 26 : perc="33" : tresh=9
					case 27 : perc="33" : tresh=9
					case else : perc="33" : tresh=round(cMediafiles.count/3,0)+1 : countB=true
					
				end select	
				
				dim margintop
				select case iPMlocation
					case 2,4 : margintop="5"
					case else : margintop="20"
				end select			
				
				medfstr=vbcrlf & "<div class=""qspagemedia qsrow"" style=""margin-bottom:20px;margin-top:" & margintop & "px"">" & vbcrlf
				medfstr=medfstr & "<div class=""qscolumn"" style=""" & padding & "flex:" & perc & "%;max-width:" & perc & "%"">" & vbcrlf
				
				dim countB, countf, lightboxSAMPLE : lightboxSAMPLE=generatePassword : countf=0
				
				for each medf in cMediafiles	
					
					if countf=tresh then
					
						if countB then tresh=tresh-1 : countB=false
						
						countf=0
						medfstr=medfstr & "</div>" & vbcrlf
						medfstr=medfstr & "<div class=""qscolumn"" style="""&padding&"flex:" & perc & "%;max-width:" & perc & "%"">" & vbcrlf
					end if
					
					countf=countf+1
				
					medfstr=medfstr & "<!--" & countf & "-->"
					medfstr=medfstr & "<div class=""" & pad4 & """><a rel=""" & lightboxSAMPLE & """ href=""" & Application("QS_CMS_C_DIRECTORY_QUICKERSITE") & "/showthumb.aspx?maxsize=1920&FSR=0&img="
					medfstr=medfstr & Application("QS_CMS_C_VIRT_DIR") & Application("QS_CMS_userfiles") & "pagemedia/" & iId & "/" & medf & """ class=""QSPPIMG"">"
					medfstr=medfstr & "<img src=""" & Application("QS_CMS_C_DIRECTORY_QUICKERSITE") & "/showthumb.aspx?maxsize="&imgSize&"&FSR=0&img="
					medfstr=medfstr & Application("QS_CMS_C_VIRT_DIR") & Application("QS_CMS_userfiles") & "pagemedia/" & iId & "/" & medf & """></a></div>" & vbcrlf
						
				next
				
				set cMediafiles=nothing
				
				medfstr=medfstr & "</div></div>" & vbcrlf	
				
				select case iPMlocation		
					
					case 2 : insertMedia="<div class=""qsrow""><div class=""qscolumn"" style="""&padding&"flex:47%;max-width:47%;padding-left:1.4rem"">" & value & "</div><div class=""qscolumn"" style="""&padding&"flex: 50%; max-width: 50%;padding-left:0px"">" & medfstr & "</div></div>"
					
					case 3 : insertMedia=value & medfstr
					
					case 4 : insertMedia="<div class=""qsrow""><div class=""qscolumn"" style="""&padding&"flex:47%;max-width:47%;padding-right:1.4rem"">" & medfstr & "</div><div class=""qscolumn"" style="""&padding&"flex: 50%; max-width: 50%"">" & value & "</div></div>"
				
					case else : insertMedia=medfstr & value	
					
				end select
				
			end if
			
		else
		
			insertMedia=value
	
		end if	
	
	end function	

end class%>
