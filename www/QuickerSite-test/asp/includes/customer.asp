
<%class cls_customer
Public resetDBConn
Public iId, sName, sURL, sBannerMenu, bEnableMainRSS, bRemoveTemplatesOnSetup,iFolderSize
Public bOnline	, dUpdatedTS	, dCreatedTS, dOnlineFrom, dResetStats, sTotalPW, bApplication, bIntranet, bCatalog, bMonitor
Public bTopHeader, sTopheader, sDescription	, siteName, siteTitle, copyRight,bUseAvatars
Public webmasterEmail	, adminPassword	, webmaster, iAvatarSize
Public keywords	, language	, hasFavicon, googleAnalytics, intranetUse	, intranetName
Public intranetPWEmail,sPopupViewmode,iLoginMode,sNLTemplate
Public intranetUseMyProfile, intranetMyProfile, intranetLogOff, bScanReferer, sDefaultRSSLink
Public PublicIconColor,PublicIconColorHover, bannerApplication, defaultTemplate, sFooter, sAlternateDomains
Public bUserFriendlyURL,bAllowStorageOutsideWWW, sLeftBanner, sRightBanner, sDatumFormat, sHeader, sLabelEditSite
Public bConsiderNewAccount,bSaveIISSuccess,overrideLANGUAGE,bPackage,sSiteSlogan,iPackageID,sAvatarBorderColor
Public sContactInfo,sHighlights,sCodebase,doSaveIIS,copyFromCustomerPath,copyFromCustomerID,sNotifValidate, saveBindings,sOrigURL,bDatabaseHack,bEnableNewsletters
Public sProp01,sProp02,sProp03,sProp04,sProp05,sProp06,sProp07,sProp08
Public sPeelURL,bPeelEnabled,sPeelImage,bPeelOINW,sPeelFlipColor,sPeelIdleSize,sPeelMOSize,sMOBBrowsers,sMOBurl,iDefaultMobileTemplate,bUseCachingForPages
Public bCookieWarning,sCWLocation,sCWNumber,sCWText,sCWAccept,sCWContinue,sCWError,sCWBackgroundColor,sCWButtonClass,sCWTextColor,sCWLinkColor,bCWUseAsNormalPP
Public sQSAccordionMain,sQSAccordionHeader,sQSAccordionContent, bCustom404, sCustom404Title, sCustom404Body,i404TemplateID,sArrowUP,bListItemPic,bShoppingCart
Public SMTPSERVER,SMTPPORT,SMTPUSERNAME,SMTPUSERPW,SENDUSING,SMTPUSESSL
'intranet
Public bAllowNewRegistrations 'nieuwe registraties toelaten?
Public bSendMailUponNewMember 'send een warningemail in geval van een nieuwe member 
Public sExplTicket 'de tekst die bovenaan de pagina komt waar je je email opgeeft (optional)
Public sExplProfile 'de tekst die boven de My Profile komt (optional)
Public sMailTicket 'de tekst die verstuurd word naar een registrant met een activatiecode
Public sLabelRegister 'de label die zegt 'Register Now'
Public sMailWelcome 'email die vertrekt naar nieuwe registrant (optional)
Public sSubjectMailTicket 'het onderwerp van de email die vertrekt naar een nieuwe registreerde
Public iDefaultStatus 'default status waarin customers komen bij nieuwe registratie
Public sWelcomeMessage
Public sEmailNewRegistrations
private p_secondAdmin, p_nmrbContacts,p_iTotalHits,p_iMaxVisits,tableFeeds
public property get sVDUrl
if isLeeg(C_VIRT_DIR) then
sVDUrl=sUrl
elseif convertStr(right(sUrl,len(C_VIRT_DIR)))=convertStr(C_VIRT_DIR) then
sVDUrl=sUrl
else
sVDUrl=sUrl & C_VIRT_DIR
end if
end property
public property get sQSUrl
if isLeeg(C_DIRECTORY_QUICKERSITE) then
sQSUrl=sUrl
elseif convertStr(right(sUrl,len(C_DIRECTORY_QUICKERSITE)))=convertStr(C_DIRECTORY_QUICKERSITE) then
sQSUrl=sUrl
else
sQSUrl=sUrl & C_DIRECTORY_QUICKERSITE
end if
end property
Private Sub Class_Initialize

bCustom404=false
sCustom404Title="File Not Found"
sCustom404Body="The file ""[404FILENAME]"" you searched for cannot be found."
i404TemplateID=0
bListItemPic=false
bShoppingCart=false
bUseCachingForPages=false
iId=null
sPopupViewmode="1"
bOnline=true
sURL="http://"
p_iTotalHits=null
p_iMaxVisits=null
iPackageID=null
sLabelEditSite="Edit Site"
p_nmrbContacts=null
bSaveIISSuccess=false
bCookieWarning=false
bConsiderNewAccount=true
bPackage=false
doSaveIIS=false
saveBindings=false
sContactInfo="Contact Info"
sHighlights="Highlights"
bDatabaseHack=true
bUseAvatars=false
bMonitor=false
bEnableNewsletters=false
bRemoveTemplatesOnSetup=true
bPeelEnabled=false
bPeelOINW=false
sPeelFlipColor=1
sPeelIdleSize=70
sPeelMOSize=300
resetDBConn=true
end sub
public function getHomePageObj
dim rs
set rs=db.execute("select iId from tblPage ")
set getHomePageObj=new cls_page
getHomePageObj.pickByCodeNOCID("")
end function
Public Function Pick(id)
On Error Resume Next
if isNumeriek(id) then
dim sql, RS
sql = "select * from tblCustomer where iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iID")
bOnline	= rs("bOnline")
sName	= rs("sName")
sURL	= rs("sURL")
sBannerMenu	= rs("sBannerMenu")
dUpdatedTS	= rs("dUpdatedTS")
dCreatedTS	= rs("dCreatedTS")
dOnlineFrom	= rs("dOnlineFrom")
dResetStats	= rs("dResetStats")
sTotalPW	= rs("sTotalPW")
bApplication	= rs("bApplication")
bIntranet	= rs("bIntranet")
bCatalog	= rs("bCatalog")
bTopHeader	= rs("bTopHeader")
sTopheader	= rs("sTopheader")
sDescription	= rs("sDescription")
siteName	= rs("siteName")
siteTitle	= rs("siteTitle")
copyRight	= rs("copyRight")
webmasterEmail	= rs("webmasterEmail")
adminPassword	= rs("adminPassword")
webmaster	= rs("webmaster")
keywords	= rs("keywords")
language	= rs("language")
hasFavicon	= rs("hasFavicon")
googleAnalytics	= rs("googleAnalytics")
intranetUse	= rs("intranetUse")
intranetName	= rs("intranetName")
intranetPWEmail	= rs("intranetPWEmail")
intranetUseMyProfile	= rs("intranetUseMyProfile")
intranetMyProfile	= rs("intranetMyProfile")
intranetLogOff	= rs("intranetLogOff")
publicIconColor	= MYQS_IconColor
publicIconColorHover	= MYQS_IconHoverColor
bannerApplication	= rs("bannerApplication")
defaultTemplate	= rs("defaultTemplate")
sFooter	= rs("sFooter")
bScanReferer	= rs("bScanReferer")
sAlternateDomains	= rs("sAlternateDomains")
sDefaultRSSLink	= rs("sDefaultRSSLink")
bUserFriendlyURL	= rs("bUserFriendlyURL")
bAllowStorageOutsideWWW	= rs("bAllowStorageOutsideWWW")
sLeftBanner	= rs("sLeftBanner")
sRightBanner	= rs("sRightBanner")
iDefaultStatus	= rs("iDefaultStatus")
sWelcomeMessage	= rs("sWelcomeMessage")
sDatumFormat	= rs("sDatumFormat")
sHeader	= rs("sHeader")
bAllowNewRegistrations	= rs("bAllowNewRegistrations")
bSendMailUponNewMember	= rs("bSendMailUponNewMember")
sExplTicket	= rs("sExplTicket")
sMailTicket	= rs("sMailTicket")
sLabelRegister	= rs("sLabelRegister")
sExplProfile	= rs("sExplProfile")
sMailWelcome	= rs("sMailWelcome")
sSubjectMailTicket	= rs("sSubjectMailTicket")
sEmailNewRegistrations	= rs("sEmailNewRegistrations")
bEnableMainRSS	= rs("bEnableMainRSS")
sLabelEditSite	= rs("sLabelEditSite")
sSiteSlogan	= rs("sSiteSlogan")
iPackageID	= rs("iPackageID")
sContactInfo	= rs("sContactInfo")
sHighlights	= rs("sHighlights")
sNotifValidate	= rs("sNotifValidate")
bMonitor	= rs("bMonitor")
bEnableNewsletters	= convertBool(rs("bEnableNewsletters"))
sPopupViewmode	= rs("sPopupViewmode")
iFolderSize	= rs("iFolderSize")
bUseAvatars	= rs("bUseAvatars")
iAvatarSize	= rs("iAvatarSize")
iLoginMode	= rs("iLoginMode")
sAvatarBorderColor	= rs("sAvatarBorderColor")
sProp01	= rs("sProp01")
sProp02	= rs("sProp02")
sProp03	= rs("sProp03")
sProp04	= rs("sProp04")
sProp05	= rs("sProp05")
sProp06	= rs("sProp06")
sProp07	= rs("sProp07")
sProp08	= rs("sProp08")
bPeelEnabled	= rs("bPeelEnabled")
sPeelURL	= rs("sPeelURL")
sPeelImage	= rs("sPeelImage")
bPeelOINW	= rs("bPeelOINW")
sPeelFlipColor	= rs("sPeelFlipColor")
sPeelIdleSize	= rs("sPeelIdleSize")
sPeelMOSize	= rs("sPeelMOSize")
sNLTemplate	= rs("sNLTemplate")
sMOBBrowsers	= rs("sMOBBrowsers")
sMOBUrl	= rs("sMOBUrl")
iDefaultMobileTemplate	= rs("iDefaultMobileTemplate")
bUseCachingForPages	= rs("bUseCachingForPages")
bCookieWarning	= rs("bCookieWarning")
sCWLocation	= rs("sCWLocation")
sCWNumber	= rs("sCWNumber")
sCWText	= rs("sCWText")
sCWAccept	= rs("sCWAccept")
sCWError	= rs("sCWError")
sCWContinue	= rs("sCWContinue")
sCWBackgroundColor	= rs("sCWBackgroundColor")
sCWButtonClass	= rs("sCWButtonClass")
sCWTextColor	= rs("sCWTextColor")
sCWLinkColor	= rs("sCWLinkColor")
bCWUseAsNormalPP	= rs("bCWUseAsNormalPP")
sQSAccordionMain	= rs("sQSAccordionMain")
sQSAccordionHeader	= rs("sQSAccordionHeader")
sQSAccordionContent	= rs("sQSAccordionContent")
SMTPSERVER	= rs("C_SMTPSERVER")
SMTPPORT	= rs("C_SMTPPORT")
SMTPUSERNAME	= rs("C_SMTPUSERNAME")
SMTPUSERPW	= rs("C_SMTPUSERPW")
SENDUSING	= rs("C_SENDUSING")
SMTPUSESSL=rs("C_SMTPUSESSL")
bCustom404=rs("bCustom404")
sCustom404Title=rs("sCustom404Title")
sCustom404Body=rs("sCustom404Body")
i404TemplateID=rs("i404TemplateID")
sArrowUP=rs("sArrowUP")
bListItemPic=rs("bListItemPic")
bShoppingCart=rs("bShoppingCart")

if isLeeg(sQSAccordionMain) then sQSAccordionMain="font-family: inherit; " & vbcrlf & "font-size: 1em;"
if isLeeg(sQSAccordionHeader) then sQSAccordionHeader="background-color: #DDDDDD;" & vbcrlf & "color: #000000;" & vbcrlf & "font-weight: 700;"
if isLeeg(sQSAccordionContent) then sQSAccordionContent="background-color: #FFFFFF;" & vbcrlf & "color: #000000;"

if request.querystring("QSbcwp")=1 then
bCookieWarning=true
end if
end if
set RS = nothing
end if

'cache only for 1 hour
if convertBool(bUseCachingForPages) then
if convertGetal(application("oneHourCache"))<>convertGetal(hour(time())) then
application("oneHourCache")=convertGetal(hour(time()))
clearPageCache()
end if
end if

dumpError "Pick Customer",err
On Error Goto 0
end function
Public Function Check()
Check = true
if isLeeg(sName) then
check=false
message.AddError("err_mandatory")
end if
if len(sUrl) <= 7 then
    check=false
message.AddError("err_url")	    
end if
if intranetUse then
if isLeeg(intranetName) then
check=false
message.AddError("err_mandatory")
end if
if not isLeeg(sNotifValidate) then
sNotifValidate=trim(sNotifValidate)
if not CheckEmailSyntax(sNotifValidate) then
message.AddError("err_email")
check=false
end if
end if
if intranetUseMyProfile then
if isLeeg(intranetMyProfile) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(intranetLogOff) then
check=false
message.AddError("err_mandatory")
end if
if bAllowNewRegistrations then
if isleeg(iDefaultStatus) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(sLabelRegister) then
check=false
message.AddError("err_mandatory")
end if
if bSendMailUponNewMember then
if not isLeeg(sEmailNewRegistrations) then
if not CheckEmailSyntax(sEmailNewRegistrations) then
message.AddError("err_email")
check=false
end if
else
check=false
message.AddError("err_mandatory")
end if
end if
end if
end if
end if
End Function
Public Function Save
if check() then
save=true
else
save=false
exit function
end if
if left (lcase(sUrl),4)<>"http" and left (lcase(sUrl),3)<>"ftp" then
sUrl="http://" & sUrl
end if
if right(sURL,1)="/" then
sURL=left(sURL,len(sURL)-1)
end if
sURL=trim(lcase(sURL))
bScanReferer	= convertBool(bScanReferer)
bEnableMainRSS	= convertBool(bEnableMainRSS)
bAllowStorageOutsideWWW	= convertBool(bAllowStorageOutsideWWW)
bApplication	= convertBool(bApplication)
bMonitor	= convertBool(bMonitor)
bEnableNewsletters	= convertBool(bEnableNewsletters)
bUseCachingForPages	= convertBool(bUseCachingForPages)
bCookieWarning	= convertBool(bCookieWarning)
sFooter=removeEmptyP(sFooter)
sTopheader=removeEmptyP(sTopheader)
if resetDBConn then
'set db=nothing
'set db=new cls_database
end if
dim rs
set rs = db.GetDynamicRS
dim newRecord
newRecord=false
if isLeeg(iId) then
if bConsiderNewAccount then 
newRecord=true
initDesign()
end if
rs.Open "select * from tblCustomer where 1=2"
rs.AddNew
rs("dCreatedTS")=now()
dResetStats=now()
if isLeeg(sAlternateDomains) then
if instr(sUrl,"www.")<>0 then
sAlternateDomains=replace(sUrl,"www.","",1,-1,1)
else
sAlternateDomains=replace(sUrl,"http://","http://www.",1,-1,1)
end if
end if
else
rs.Open "select * from tblCustomer where iId="& iId
end if
rs("sName")	= sName
rs("sURL")	= sURL
rs("sBannerMenu")	= sBannerMenu
rs("bOnline")	= bOnline
rs("dOnlineFrom")	= dOnlineFrom
rs("dUpdatedTS")	= now()
rs("dResetStats")	= dResetStats
rs("sTotalPW")	= sTotalPW
rs("bApplication")	= bApplication
rs("bIntranet")	= bIntranet
rs("bCatalog")	= bCatalog
rs("bTopHeader")	= bTopHeader
rs("sTopheader")	= sTopheader
rs("sDescription")	= sDescription
rs("siteName")	= siteName
rs("siteTitle")	= siteTitle
rs("copyRight")	= copyRight
rs("webmasterEmail")	= webmasterEmail
rs("adminPassword")	= adminPassword
rs("webmaster")	= webmaster
rs("keywords")	= keywords
rs("language")	= language
rs("hasFavicon")	= hasFavicon
rs("googleAnalytics")	= googleAnalytics
rs("intranetUse")	= intranetUse
rs("intranetName")	= intranetName
rs("intranetPWEmail")	= intranetPWEmail
rs("intranetUseMyProfile")	= intranetUseMyProfile
rs("intranetMyProfile")	= intranetMyProfile
rs("intranetLogOff")	= intranetLogOff
rs("bannerApplication")	= bannerApplication
rs("defaultTemplate")	= defaultTemplate
rs("sFooter")	= sFooter
rs("bScanReferer")	= bScanReferer
rs("sAlternateDomains")	= sAlternateDomains
rs("sDefaultRSSLink")	= sDefaultRSSLink
rs("bUserFriendlyURL")	= bUserFriendlyURL
rs("bAllowStorageOutsideWWW")	= bAllowStorageOutsideWWW
rs("sRightBanner")	= sRightBanner
rs("sLeftBanner")	= sLeftBanner
rs("sDatumFormat")	= sDatumFormat
rs("sHeader")	= sHeader
rs("bAllowNewRegistrations")	= bAllowNewRegistrations 'ok
rs("iLoginMode")	= convertGetal(iLoginMode)
rs("bSendMailUponNewMember")	= bSendMailUponNewMember 'na
rs("sExplTicket")	= sExplTicket 'ok
rs("sMailTicket")	= sMailTicket 'ok
rs("sLabelRegister")	= sLabelRegister 'ok
rs("sExplProfile")	= sExplProfile 'ok
rs("sMailWelcome")	= sMailWelcome	'na
rs("sSubjectMailTicket")	= sSubjectMailTicket 'ok
rs("iDefaultStatus")	= iDefaultStatus
rs("sWelcomeMessage")	= sWelcomeMessage
rs("sEmailNewRegistrations")	= sEmailNewRegistrations
rs("bEnableMainRSS")	= bEnableMainRSS
rs("sLabelEditSite")	= sLabelEditSite
rs("sSiteSlogan")	= sSiteSlogan
rs("iPackageID")	= iPackageID
rs("sContactInfo")	= sContactInfo
rs("sHighlights")	= sHighlights
rs("sNotifValidate")	= sNotifValidate
rs("bMonitor")	= bMonitor
rs("bEnableNewsletters")	= bEnableNewsletters
rs("bUseCachingForPages")	= bUseCachingForPages
rs("bCookieWarning")	= bCookieWarning
rs("sPopupViewmode")	= sPopupViewmode
rs("iFolderSize")	= convertGetal(iFolderSize)
rs("bUseAvatars")	= convertBool(bUseAvatars)
rs("iAvatarSize")	= iAvatarSize
rs("sAvatarBorderColor")	= sAvatarBorderColor
rs("sProp01")	= sProp01
rs("sProp02")	= sProp02
rs("sProp03")	= sProp03
rs("sProp04")	= sProp04
rs("sProp05")	= sProp05
rs("sProp06")	= sProp06
rs("sProp07")	= sProp07
rs("sProp08")	= sProp08
rs("sPeelImage")	= sPeelImage
rs("sPeelURL")	= sPeelURL
rs("bPeelEnabled")	= bPeelEnabled
rs("bPeelOINW")	= bPeelOINW
rs("sPeelFlipColor")	= sPeelFlipColor
rs("sPeelIdleSize")	= sPeelIdleSize
rs("sPeelMOSize")	= sPeelMOSize
rs("sNLTemplate")	= sNLTemplate
rs("sMOBUrl")	= sMOBUrl
rs("sMOBBrowsers")	= sMOBBrowsers
rs("iDefaultMobileTemplate")	= iDefaultMobileTemplate
rs("sCWNumber")	= convertGetal(sCWNumber)
rs("sCWLocation")	= sCWLocation
rs("sCWText")	= sCWText
rs("sCWAccept")	= sCWAccept
rs("sCWError")	= sCWError
rs("sCWContinue")	= sCWContinue
rs("sCWButtonClass")	= sCWButtonClass
rs("sCWTextColor")	= sCWTextColor
rs("sCWLinkColor")	= sCWLinkColor
rs("bCWUseAsNormalPP")	= convertBool(bCWUseAsNormalPP)
rs("sQSAccordionMain")	= sQSAccordionMain
rs("sQSAccordionHeader")	= sQSAccordionHeader
rs("sQSAccordionContent")	= sQSAccordionContent
rs("C_SMTPSERVER")	= SMTPSERVER
rs("C_SMTPPORT")	= SMTPPORT
rs("C_SMTPUSERNAME")	= SMTPUSERNAME
rs("C_SMTPUSERPW")	= SMTPUSERPW
rs("C_SENDUSING")	= SENDUSING
rs("C_SMTPUSESSL") = SMTPUSESSL

rs("bCustom404")=bCustom404
rs("sCustom404Title")=sCustom404Title
rs("sCustom404Body")=sCustom404Body
rs("i404TemplateID")=i404TemplateID
rs("sArrowUP")=sArrowUP
rs("bListItemPic")=convertBool(bListItemPic)
rs("bShoppingCart")=convertBool(bShoppingCart)

if isLeeg(sCWBackgroundColor) then sCWBackgroundColor="#000000"
if isLeeg(sCWTextColor) then sCWTextColor="#FFFFFF"
rs("sCWBackgroundColor")	= sCWBackgroundColor
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
if newRecord or doSaveIIS then
if not doSaveIIS then initialize()
on error resume next
dim iisSite
set iisSite=new cls_iisSite
iisSite.bDatabasehack=bDatabaseHack
iisSite.sHostname =lcase(replace(sURL,"http://","",1,-1,1))
iisSite.sSiteName =iisSite.sHostname
if not isLeeg(copyFromCustomerPath) then
iisSite.samplesitepath=copyFromCustomerPath
end if
if not isLeeg(copyFromCustomerID) then
iisSite.copyFromCustomerID=copyFromCustomerID
end if
if sCodebase<>"" then
iisSite.QS_path=sCodebase
end if
iisSite.otherbindings=sAlternateDomains
iisSite.sPath ="HostedCMS" & iId & "." & iisSite.sHostname 
iisSite.SiteiId=iId
if iisSite.create then 
bSaveIISSuccess=true
end if
set iisSite=nothing
on error goto 0
end if
if saveBindings then
on error resume next
dim iisSiteSaveBindings
set iisSiteSaveBindings=new cls_iisSite
'iisSiteSaveBindings.updateServerBindings sOrigURL,sUrl & vbcrlf & sAlternateDomains
set iisSiteSaveBindings=nothing
on error goto 0
end if
end function
private sub initialize()
'set db=nothing
'set db=new cls_database
'create Homepage
dim homepage
set homepage = db.GetDynamicRS
homepage.Open "select * from tblPage where 1=2"
homepage.AddNew
homepage("createdTS")	= now()
homepage("iParentid")	= null
homepage("iCustomerID")	= iId
homepage("sTitle")	= "Home"
homepage("sValue")	= "<p>Homepage</p>"
homepage("iRang")	= 1
homepage("bOnline")	= true
homepage("bHomepage")	= true
homepage("bIntranet")	= false
 	homepage("bDeleted")	= false
homepage("bContainerPage")	= false
homepage("iCatalogId")	= null
homepage("updatedTS")	= now()
if bPackage then
homepage("sUserFriendlyURL")	= "index.html"
end if
homepage.Update 
homepage.close
Set homepage = nothing
'set db=nothing
'set db=new cls_database
set homepage = db.GetDynamicRS
homepage.Open "select * from tblPage where 1=2"
homepage.AddNew
homepage("createdTS")	= now()
homepage("iParentid")	= null
homepage("iCustomerID")	= iId
homepage("sTitle")	= "Intranet"
homepage("sValue")	= "<p>Homepage Intranet</p>"
homepage("iRang")	= 1
homepage("bOnline")	= true
homepage("bHomepage")	= true
homepage("bIntranet")	= true
homepage("bDeleted")	= false
homepage("bContainerPage")	= false
homepage("iCatalogId")	= null
homepage("updatedTS")	= now()
if bPackage then
homepage("sUserFriendlyURL")	= "intranet.html"
end if
homepage.Update 
homepage.close
Set homepage = nothing
'enkele parameters opvullen
siteName	= sName
siteTitle	= sName
copyRight	= sName
sDatumFormat	= QS_dateFM_EU
bMonitor	= false
bEnableNewsletters	= false
save()
end sub
public function remove
if convertGetal(iId)<>0 then
reset()
db.execute("delete from tblMonitor where iCustomerID="&iId)
db.execute("delete from tblPage where iCustomerID="&iId)
db.execute("delete from tblCustomer where iId="&iId)
On Error Resume Next
dim iisSite
set iisSite=new cls_iisSite
iisSite.sHostname =replace(sURL,"http://","",1,-1,1)
'iisSite.remove
set iisSite=nothing
On Error Goto 0
end if
end function
public function reset()
if not isLeeg(iId) then
sBannerMenu=""
bMonitor=false
sPopupViewmode="1"
sFooter	=""
sTopheader=""
sHeader=""
bannerApplication=""
dOnlineFrom=now()
sPeelURL=""
bPeelEnabled=false
sPeelImage=""
bPeelOINW=false
dim cCat,cMail,cField
dim cCats,cMails,cFields
dim cForm,cForms
dim cFeeds,cFeed
dim cthemes,ctheme
dim cGalleries,gallery
dim cPolls,poll
dim cGuestbooks,guestbook
dim cPopups,popup
set cCats=catalogs
set cMails=mails
set cFields=contactFields(false)
set cForms=forms
set cFeeds=feeds
set cthemes=themes
set cGalleries=galleries
set cGuestbooks=guestbooks
set cPolls=polls
set cPopups=popups
'remove catalogs
for each cCat in cCats
cCats(cCat).remove
next
'remove mails
for each cMail in cMails
cMails(cMail).delete
next
'remove contactFields
for each cField in cFields
cFields(cField).remove()
next
'remove forms
for each cForm in cForms
cForms(cForm).remove()
next
'remove feeds
for each cFeed in cFeeds
cFeeds(cFeed).remove()
next
'remove themes
for each ctheme in cthemes
cthemes(ctheme).remove()
next
'remove galleries
for each gallery in cGalleries
cGalleries(gallery).remove()
next
'remove polls
for each poll in cPolls
cPolls(poll).remove()
next
'remove guestbooks
for each guestbook in cguestbooks
cguestbooks(guestbook).remove()
next
'remove popups
for each popup in cpopups
cpopups(popup).remove()
next
on error resume next
db.execute("delete from tblNewsletterCategorySubscriber where iCustomerID="&iId)
db.execute("delete from tblPoll where iCustomerID="&iId)
db.execute("delete from tblFeed where iCustomerID="&iId)
db.execute("delete from tblGallery where iCustomerID="&iId)
db.execute("delete from tblGuestbook where iCustomerID="&iId)
db.execute("delete from tblMonitor where iCustomerID="&iId)
db.execute("delete from tblPage where iCustomerID="&iId)
db.execute("delete from tblConstant where iCustomerID="&iId)
db.execute("delete from tblContact where iCustomerID="&iId)
db.execute("delete from tblContactRegistration where iCustomerID="&iId)
db.execute("delete from tblCustomerIntranetMessage where iCustomerID="&iId)
db.execute("delete from tblSession where iCustomerID="&iId)
on error goto 0
if bRemoveTemplatesOnSetup then
db.execute("delete from tblTemplate where iCustomerID="&iId)
end if
db.execute("delete from tblSecondAdmin where iCustomerID="&iId)
initialize()
initDesign()
save()
end if
end function
public sub resetStats()
dim rs
set rs=db.execute("update tblPage set iHits=0, iVisitors=0, iHitsRSS=0 where iCustomerID="&iID)
set rs=nothing
dResetStats=now()
save()
end sub
Public property get iTotalHits
if isNull(p_iTotalHits) then
dim rs
set rs=db.execute("select sum(iHits) from tblPage where bOnline="&getSQLBoolean(true)&" and iCustomerID=" & iID)
p_iTotalHits=convertGetal(rs(0))
set rs=nothing
end if
iTotalHits=p_iTotalHits
end property
Public property get iMaxVisits
if isNull(p_iMaxVisits) then
dim rs
set rs=db.execute("select max(iVisitors) from tblPage where bOnline="&getSQLBoolean(true)&" and iCustomerID=" & iID)
p_iMaxVisits=convertGetal(rs(0))
set rs=nothing
end if
iMaxVisits=p_iMaxVisits
end property
Public function applyTotalPW(value)
if len(value)<3 then
applyTotalPW=false
message.AddError("err_pw")
else
sTotalPW=lcase(convertStr(value))
save()
dim rs
set rs=db.execute("update tblPage set sPw='"& value & "' where (bIntranet is null or bIntranet="&getSQLBoolean(false)&") and iCustomerID="& iID)
set rs=nothing
applyTotalPW=true
end if
end function
Public function removeTotalPW()
removeTotalPW=true
sTotalPW=""
save()
dim rs
set rs=db.execute("update tblPage set sPw='' where iCustomerID="& iID)
set rs=nothing
end function
Public property get aantalDagen
aantalDagen=convertGetal(dateDiff("d",dResetStats,date))
if aantalDagen=0 then aantalDagen=1
end property
public function contactFields(searchFieldsOnly)
set contactFields=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, contactField, sql
sql="select iId from tblContactField where iCustomerID="& iId
if searchFieldsOnly then
sql=sql&" and bSearchField="&getSQLBoolean(true)&" "
end if
sql=sql&" order by iRang asc"
set rs=db.execute(sql)
while not rs.eof
set contactField=new cls_contactField
contactField.pick(rs(0))
contactFields.Add contactField.iID,contactField
set contactField=nothing
rs.movenext
wend
set rs=nothing
end if
end function
public function mails
on error resume next
set mails=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, mail
set rs=db.execute("select iId from tblMail where iCustomerID="& iId & " order by dDateSent desc")
while not rs.eof
set mail=new cls_mail
mail.pick(convertGetal(rs(0)))
mails.Add mail.iId, mail
set mail=nothing
rs.movenext
wend
set rs=nothing
end if
end function
public function catalogs
on error resume next
set catalogs=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, catalog, sql
sql="select iId from tblCatalog where iCustomerID="& iId & " order by sName"
set rs=db.execute(sql)
while not rs.eof
set catalog=new cls_catalog
catalog.pick(rs(0))
catalogs.Add catalog.iID,catalog
set catalog=nothing
rs.movenext
wend
set rs=nothing
end if
end function
Public Function showSelectedCatalog(mode, selected)
dim copyCats
set copyCats=catalogs
Select Case mode
Case "single"
showSelectedCatalog = copyCats(convertGetal(selected))
Case "option"
Dim key
For each key in copyCats
showSelectedCatalog = showSelectedCatalog & "<option value='" & encrypt(key) & "'"
If convertStr(encrypt(selected)) = convertStr(encrypt(key)) Then
showSelectedCatalog = showSelectedCatalog & " selected"
End If
showSelectedCatalog = showSelectedCatalog & ">" & copyCats(key).sName & "</option>"
Next
End Select
End Function
Public Function showSelectedNewsletter(mode, selected)
dim copyCats
set copyCats=newsletters
Select Case mode
Case "single"
showSelectedNewsletter = copyCats(convertGetal(selected))
Case "option"
Dim key
For each key in copyCats
showSelectedNewsletter = showSelectedNewsletter & "<option value='" & key & "'"
If convertStr(selected) = convertStr(key) Then
showSelectedNewsletter = showSelectedNewsletter & " selected"
End If
showSelectedNewsletter = showSelectedNewsletter & ">" & copyCats(key).sName & "</option>"
Next
End Select
End Function
public function forms
on error resume next
set forms=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, form, sql
sql="select iId from tblForm where iCustomerID="& iId & " order by sName"
set rs=db.execute(sql)
while not rs.eof
set form=new cls_form
form.pick(rs(0))
forms.Add form.iID,form
set form=nothing
rs.movenext
wend
set rs=nothing
end if
on error goto 0
end function
public function feeds
on error resume next
set feeds=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, feed, sql
sql="select iId from tblFeed where iCustomerID="& iId & " order by sName"
set rs=db.execute(sql)
while not rs.eof
set feed=new cls_feed
feed.pick(rs(0))
feeds.Add feed.iID,feed
set feed=nothing
rs.movenext
wend
set rs=nothing
end if
on error goto 0
end function
public function polls
on error resume next
set polls=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, poll, sql
sql="select iId from tblPoll where iCustomerID="& iId & " order by sCode"
set rs=db.execute(sql)
while not rs.eof
set poll=new cls_poll
poll.pick(rs(0))
polls.Add poll.iID,poll
set poll=nothing
rs.movenext
wend
set rs=nothing
end if
end function
public function popups
on error resume next
set popups=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, popup, sql
sql="select iId from tblpopup where iCustomerID="& iId & " order by sName"
set rs=db.execute(sql)
while not rs.eof
set popup=new cls_popup
popup.pick(rs(0))
popups.Add popup.iID,popup
set popup=nothing
rs.movenext
wend
set rs=nothing
end if
end function
public function newsletters
on error resume next
set newsletters=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, newsletter, sql
sql="select iId from tblNewsletter where iCustomerID="& iId & " order by iId desc"
set rs=db.execute(sql)
while not rs.eof
set newsletter=new cls_newsletter
newsletter.pick(rs(0))
newsletters.Add newsletter.iID,newsletter
set newsletter=nothing
rs.movenext
wend
set rs=nothing
end if
end function
public function newsletterCategories
on error resume next
set newsletterCategories=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, newsletterCAT, sql
sql="select iId from tblNewsletterCategory where iCustomerID="& iId & " order by sName"
set rs=db.execute(sql)
while not rs.eof
set newsletterCAT=new cls_newsletterCategory
newsletterCAT.pick(rs(0))
newsletterCategories.Add newsletterCAT.iID,newsletterCAT
set newsletterCAT=nothing
rs.movenext
wend
set rs=nothing
end if
end function
public property get bCanImportSubscribers
on error resume next
dim rs
set rs=db.execute("select count(*) from tblNewsletterCategory where iCustomerID="& iId)
bCanImportSubscribers=convertGetal(rs(0))>0
set rs=nothing
on error goto 0
end property
public function guestbooks
on error resume next
set guestbooks=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, guestbook, sql
sql="select iId from tblguestbook where iCustomerID="& iId & " order by sCode"
set rs=db.execute(sql)
while not rs.eof
set guestbook=new cls_guestbook
guestbook.pick(rs(0))
guestbooks.Add guestbook.iID,guestbook
set guestbook=nothing
rs.movenext
wend
set rs=nothing
end if
end function
public function galleries
on error resume next
set galleries=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, gallery, sql
sql="select iId from tblGallery where iCustomerID="& iId & " order by sName"
set rs=db.execute(sql)
while not rs.eof
set gallery=new cls_gallery
gallery.pick(rs(0))
galleries.Add gallery.iID,gallery
set gallery=nothing
rs.movenext
wend
set rs=nothing
end if
on error goto 0
end function
public function templates
on error resume next
set templates=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, template, sql
sql="select iId from tbltemplate where iCustomerID="& iId & " order by sName"
set rs=db.execute(sql)
while not rs.eof
set template=new cls_template
template.pick(rs(0))
if not 	templates.Exists (template.iID) then
templates.Add template.iID,template
end if
set template=nothing
rs.movenext
wend
set rs=nothing
end if
end function
Public Function showSelectedForm(mode, selected)
dim copyForms
set copyForms=forms
Select Case mode
Case "single"
showSelectedForm = copyForms(convertGetal(selected))
Case "option"
Dim key
For each key in copyForms
showSelectedForm = showSelectedForm & "<option value='" & encrypt(key) & "'"
If convertStr(encrypt(selected)) = convertStr(encrypt(key)) Then
showSelectedForm = showSelectedForm & " selected"
End If
showSelectedForm = showSelectedForm & ">" & copyForms(key).sName & "</option>"
Next
End Select
End Function
Public Function showSelectedfeed(mode, selected)
dim copyfeeds
set copyfeeds=feeds
Select Case mode
Case "single"
showSelectedfeed = copyfeeds(convertGetal(selected))
Case "option"
Dim key
For each key in copyfeeds
showSelectedfeed = showSelectedfeed & "<option value='" & encrypt(key) & "'"
If convertStr(encrypt(selected)) = convertStr(encrypt(key)) Then
showSelectedfeed = showSelectedfeed & " selected"
End If
showSelectedfeed = showSelectedfeed & ">" & copyfeeds(key).sName & "</option>"
Next
End Select
End Function
Public Function showSelectedtemplate(mode, selected)
dim copytemplates
set copytemplates=templates
Select Case mode
Case "single"
showSelectedtemplate = copytemplates(convertGetal(selected))
Case "option"
Dim key
For each key in copytemplates
showSelectedtemplate = showSelectedtemplate & "<option value='" & encrypt(key) & "'"
If convertStr(encrypt(selected)) = convertStr(encrypt(key)) Then
showSelectedtemplate = showSelectedtemplate & " selected"
End If
showSelectedtemplate = showSelectedtemplate & ">" & copytemplates(key).sName & "</option>"
Next
End Select
End Function
public property get nmbrParentMenus
if isNumeriek(iId) then
dim rs,sql
sql="select count(*) from tblPage where iCustomerID="& iId & " and (iListPageId is null or iListPageId=0) "
sql=sql&" and bLossePagina="&getSQLBoolean(false)&" and bOnline="&getSQLBoolean(true)&" and (iParentid is null or iParentid=0 and bDeleted="&getSQLBoolean(false)&" "
if not logon.authenticatedIntranet then
sql=sql&" and bIntranet="&getSQLBoolean(false)&" "
end if
set rs=db.execute(sql)
nmbrParentMenus=clng(rs(0))
set rs=nothing
else
nmbrParentMenus=0
end if
if intranetUse and not logon.authenticatedIntranet then nmbrParentMenus=nmbrParentMenus+1
if intranetUse and intranetUseMyProfile then nmbrParentMenus=nmbrParentMenus+1
if intranetUse and not isleeg(intranetLogOff) then nmbrParentMenus=nmbrParentMenus+1
'Response.Write nmbrParentMenus
end property
public sub getIntranetAdminRequestValues()
intranetUse	= convertBool(Request.Form ("intranetUse"))
intranetName	= Request.Form ("intranetName")
intranetUseMyProfile	= convertBool(Request.Form ("intranetUseMyProfile"))
bAllowNewRegistrations	= convertBool(Request.Form ("bAllowNewRegistrations"))
sEmailNewRegistrations	= Request.Form ("sEmailNewRegistrations")
bSendMailUponNewMember	= convertBool(Request.Form ("bSendMailUponNewMember"))
iDefaultStatus	= convertGetal(Request.Form ("iDefaultStatus"))
sLabelRegister	= Request.Form ("sLabelRegister")
sLabelEditSite	= Request.Form ("sLabelEditSite")
intranetMyProfile	= Request.Form ("intranetMyProfile")
intranetLogOff	= Request.Form ("intranetLogOff")
sNotifValidate	= Request.Form ("sNotifValidate")
bUseAvatars	= convertBool(request.form("bUseAvatars"))
iAvatarSize	= request.form("iAvatarSize")
iLoginMode	= request.form("iLoginMode")
sAvatarBorderColor	= request.form("sAvatarBorderColor")
if bSendMailUponNewMember and isLeeg(sEmailNewRegistrations) then
sEmailNewRegistrations=customer.webmasteremail
end if
end sub
public function defaultTemplateObj
set defaultTemplateObj=new cls_template
defaultTemplateObj.pick(defaultTemplate)
end function
public sub initDesign
'general
if isLeeg(overrideLANGUAGE) then
language=1 'english
else
language=overrideLANGUAGE
end if
hasFavicon=false
intranetUse=true
intranetName="Intranet"
intranetUseMyProfile=true
intranetMyProfile="My Profile"
intranetLogOff="Logoff"
bAllowNewRegistrations=false
bSendMailUponNewMember=false
iDefaultStatus=cs_read 'veilige scenario
sLabelRegister="Not a member yet? Register now!"
bIntranet=true
bCatalog=true
bTopHeader=true
sRightBanner=""
sLeftBanner=""
sNotifValidate=""
end sub
public function constants
set constants=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, constant, sql
sql="select iId from tblConstant where iCustomerID="& iId & " order by sconstant"
set rs=db.execute(sql)
while not rs.eof
set constant=new cls_constant
constant.pick(rs(0))
constants.Add constant.iID,constant
set constant=nothing
rs.movenext
wend
set rs=nothing
end if
end function
public sub cacheConstants
dim copyconstants, cconstantKey, iRunner
iRunner=0
set copyconstants=constants
redim arrconstants(2,copyconstants.count)
for each cconstantKey in copyconstants
arrconstants(0,iRunner)=copyconstants(cconstantKey).sconstant
if copyconstants(cconstantKey).bOnline then
arrconstants(1,iRunner)=copyconstants(cconstantKey).sValue
arrconstants(2,iRunner)=""
if copyconstants(cconstantKey).iType=QS_VBScript then
arrconstants(1,iRunner)=arrconstants(1,iRunner) & QS_VBScriptIdentifier & copyconstants(cconstantKey).sParameters
arrconstants(2,iRunner)=copyconstants(cconstantKey).sGlobal
end if
else
arrconstants(1,iRunner)=""
arrconstants(2,iRunner)=""
end if
iRunner=iRunner+1
next 
set copyconstants=nothing
application(QS_CMS_arrconstants)=arrconstants
application(QS_CMS_constantsreloaded)="true"
end sub
public function cacheFeeds
dim copyfeeds, cfeedKey, iRunner
iRunner=0
set copyfeeds=feeds
for each cfeedKey in copyfeeds
if isLeeg(copyfeeds(cfeedKey).sCode) then
copyfeeds.remove(cfeedKey)
end if
next 
redim arrfeeds(1,copyfeeds.count)
for each cfeedKey in copyfeeds
arrfeeds(0,iRunner)=copyfeeds(cfeedKey).iId
arrfeeds(1,iRunner)=copyfeeds(cfeedKey).sCode
iRunner=iRunner+1
next 
set copyfeeds=nothing
application(QS_CMS_arrfeeds)=arrfeeds
application(QS_CMS_feedsreloaded)="true"
end function
public function cacheGalleries
dim copygalleries, cGalleryKey, iRunner
iRunner=0
set copygalleries=galleries
for each cGalleryKey in copygalleries
if isLeeg(copygalleries(cGalleryKey).sCode) then
copygalleries.remove(cGalleryKey)
end if
next 
redim arrgalleries(1,copygalleries.count)
for each cGalleryKey in copygalleries
arrgalleries(0,iRunner)=copygalleries(cGalleryKey).iId
arrgalleries(1,iRunner)=copygalleries(cGalleryKey).sCode
iRunner=iRunner+1
next 
set copygalleries=nothing
application(QS_CMS_arrgalleries)=arrgalleries
application(QS_CMS_galleriesreloaded)="true"
end function
public function cachethemes
dim copythemes, cthemeKey, iRunner
iRunner=0
set copythemes=themes
for each cthemeKey in copythemes
if isLeeg(copythemes(cthemeKey).sCode) then
copythemes.remove(cthemeKey)
end if
next 
redim arrthemes(1,copythemes.count)
for each cthemeKey in copythemes
arrthemes(0,iRunner)=copythemes(cthemeKey).iId
arrthemes(1,iRunner)=copythemes(cthemeKey).sCode
iRunner=iRunner+1
next 
set copythemes=nothing
application(QS_CMS_arrthemes)=arrthemes
application(QS_CMS_themesreloaded)="true"
end function
public function saveAdminPW(sAdminPW)
if not isleeg(sAdminPW) then
if sha256(sAdminPW)=sha256(QS_defaultPW) and not devVersion() then
saveAdminPW=false
message.AddError("err_backsitepw")
else
adminPassword=sha256(sAdminPW)
saveAdminPW=Save
logon.logon customer.adminPassword
end if
else
saveAdminPW=false
message.AddError("err_mandatory")
end if
end function
public function tickets
set tickets=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, ticket, sql
sql="select iId from tblContactRegistration where iCustomerID="& iId & " order by dCreatedTS desc"
set rs=db.execute(sql)
while not rs.eof
set ticket=new cls_ticket
ticket.pick(rs(0))
tickets.Add ticket.iID,ticket
set ticket=nothing
rs.movenext
wend
set rs=nothing
end if
end function
public function intranetMessages
set intranetMessages=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, iMess, sql
sql="select iStatus from tblCustomerIntranetMessage where iCustomerID="& cId
set rs=db.execute(sql)
while not rs.eof
set iMess=new cls_customerImess
iMess.pick(rs(0))
intranetMessages.Add iMess.iStatus,iMess
set iMess=nothing
rs.movenext
wend
set csShortList=new cls_contactStatusList
set csShortList=csShortList.list
dim key
for each key in csShortList
if not intranetMessages.Exists (key) then
set iMess=new cls_customerImess
iMess.iStatus=key
intranetMessages.Add iMess.iStatus,iMess
end if 
next 
set rs=nothing
end if
end function
public property get secondAdmin
if not isObject(p_secondAdmin) then
set p_secondAdmin=new cls_secondAdmin
end if
set secondAdmin=p_secondAdmin
end property
public function themes
set themes=server.CreateObject ("scripting.dictionary")
if isNumeriek(iId) then
dim rs, theme, sql
sql="select iId from tblTheme where iCustomerID="& iId & " order by sName"
set rs=db.execute(sql)
while not rs.eof
set theme=new cls_theme
theme.pick(rs(0))
if not themes.exists(theme.iId) then themes.Add theme.iID,theme
set theme=nothing
rs.movenext
wend
set rs=nothing
end if
end function
Public Function showSelectedTheme(mode, selected)
dim copythemes
set copythemes=themes
Select Case mode
Case "single"
showSelectedTheme = copythemes(convertGetal(selected))
Case "option"
Dim key
For each key in copythemes
showSelectedTheme = showSelectedTheme & "<option value='" & encrypt(key) & "'"
If convertStr(encrypt(selected)) = convertStr(encrypt(key)) Then
showSelectedTheme = showSelectedTheme & " selected"
End If
showSelectedTheme = showSelectedTheme & ">" & copythemes(key).sName & "</option>"
Next
End Select
End Function
public property get hasManyContacts
if nmrbContacts>manyContacts then
hasManyContacts=true
else
hasManyContacts=false
end if
end property
public property get nmrbContacts
if isNull(p_nmrbContacts) then
dim rs
set rs=db.execute("select count(*) from tblContact where iCustomerID="& cId)
p_nmrbContacts=clng(rs(0))
set rs=nothing
end if
nmrbContacts=p_nmrbContacts
end property
public function copy()
if convertGetal(iId)=0 then exit function
dim oldID,newID
oldID=iId
doSaveIIS=true
iId=null
dOnlineFrom=now()
bConsiderNewAccount=false
adminPassword	= sha256(QS_defaultPW)
save() 
newID=iId
'copy default template
defaultTemplateObj.copyToCustomerID iId,true
'copy all pages
dim rs,pageObj
set rs=db.execute("select iId from tblPage where bDeleted="&getSQLBoolean(false)&" and iCustomerID=" & oldID)
'newID's - oldID's
dim ids,oldPageID
set ids=server.CreateObject ("scripting.dictionary")
set tableFeeds=server.CreateObject ("scripting.dictionary")
while not rs.eof
set pageObj=new cls_page
pageObj.overruleCID=iId
pageObj.pick(rs(0))
pageObj.copyToCustomerID(iId)
ids.Add convertGetal(rs(0)),pageObj.iId
set pageObj=nothing
rs.movenext
wend 
set rs=nothing
set rs=db.execute("select iId from tblPage where bDeleted="&getSQLBoolean(false)&" and iCustomerID=" & iId)
while not rs.eof
set pageObj=new cls_page
pageObj.overruleCID=iId
pageObj.pick(rs(0))
if ids.Exists (pageObj.iListPageID) then
pageObj.iListPageID=ids(pageObj.iListPageID)
end if
if ids.Exists (pageObj.iParentID) then
pageObj.iParentID=ids(pageObj.iParentID)
end if
'nieuwe template ophalen
if convertGetal(pageObj.iTemplateID) <>0 then
if convertGetal(defaultTemplate)<>convertGetal(pageObj.iTemplateID) then
'create template
dim PageTemplate
set PageTemplate=new cls_template
PageTemplate.overruleCID=iId
PageTemplate.pick(pageObj.iTemplateID)
PageTemplate.copyToCustomerID iId,false
pageObj.iTemplateID=PageTemplate.iId
set PageTemplate=nothing
end if
end if
pageObj.save()
set pageObj=nothing
rs.movenext
wend 
set rs=nothing
'copyAllTables
copyTable "tblCustomerIntranetMessage","iCustomerID",oldID,iId
copyTable "tblConstant","iCustomerID",oldID,iId
copyTable "tblContactField","iCustomerID",oldID,iId
copyTable "tblContactRegistration","iCustomerID",oldID,iId
copyTable "tblGallery","iCustomerID",oldID,iId
copyTable "tblMail","iCustomerID",oldID,iId
copyTable "tblSession","iCustomerID",oldID,iId
copyTable "tblSecondAdmin","iCustomerID",oldID,iId
copyTable "tblFeed","iCustomerID",oldID,iId
copyTable "tblGuestbook","iCustomerID",oldID,iId
copyTable "tblPoll","iCustomerID",oldID,iId
copyTable "tblPopup","iCustomerID",oldID,iId
'update tblFeed
dim fId
for each fId in tableFeeds
db.execute("update tblPage set iFeedID="& tableFeeds(fId) & " where iFeedID="& fId & " and iCustomerID="& iId)
next
'################################################
'copy catalogs
dim cCatalogs,cCatKey
iId=oldID
set cCatalogs=catalogs
for each cCatKey in cCatalogs
cCatalogs(cCatKey).copyToCustomer(newID)
next
set cCatalogs=nothing
iId=newID
'################################################
'################################################
'copy forms
dim cForms,cFormKey,cFormObj
iId=oldID
set cForms=forms
for each cFormKey in cForms
set cFormObj=new cls_form
cFormObj.overruleCID=iId
cFormObj.pick(cFormKey)
cFormObj.copyToCustomer(newID)
set cFormObj=nothing
next
iId=newID
set cForms=nothing
'################################################
'update email addresses
set rs=db.execute("update tblGuestBook set sEmail='" & cleanup(webmasterEmail) & "' where iCustomerID=" & iId)
set rs=nothing
set rs=db.execute("update tblForm set sTo='" & cleanup(webmasterEmail) & "' where iCustomerID=" & iId)
set rs=nothing
end function
public sub copyTable (tablename,custCol,oldID,newID)
'set db=nothing
'set db=new cls_database
'hier hebben we de nieuwe iId
dim sql, rs, rsNew, field
set rs=db.getDynamicRS
rs.open "select * from " & tablename & " where " & custCol & "=" & oldID
dim dbCopy
while not rs.eof
set dbCopy=new cls_database
set rsNew=dbCopy.getDynamicRS
rsNew.open "select * from " & tablename & " where 1=2"
rsNew.AddNew()
for each field in rs.fields
if field.name<>"iId" then
if lcase(field.name) = lcase(custCol) then
rsNew(field.name)	= newID
else
rsNew(field.name)	= rs(field.name)
end if
end if
next
rsNew.update()
select case tablename
case "tblFeed"
tableFeeds.Add convertGetal(rs("iId")),convertGetal(rsNew("iId"))
end select
rsNew.close
set rsNew=nothing
rs.movenext
set dbCopy=nothing
wend 
set rs=nothing
end sub
public function pagesTobeValidated
on error resume next
set pagesTobeValidated=server.CreateObject ("scripting.dictionary")
'exit function
dim rs,sql,keyID,keyTitle
sql="select iId,sTitle from tblPage where " & sqlCustId & " AND bDeleted="&getSQLBoolean(false)
sql=sql&" AND ((sValueToBeValidated is not null) "
sql=sql&" OR (sTitleToBeValidated  is not null))"
set rs=db.execute(sql)
while not rs.eof
keyID=rs(0)
keyTitle=rs(1)
pagesTobeValidated.Add keyID,keyTitle
rs.movenext
wend 
set rs=nothing
if err.number<>0 then
set pagesTobeValidated=server.CreateObject ("scripting.dictionary")
end if
on error goto 0
end function
public sub addLog()
on error resume next
dim oSRS
set oSRS=db.getDynamicRS
oSRS.open("SELECT * FROM tblMonitor where 1=2")
oSRS.AddNew
oSRS("sDetail")=getVisitorDetails
oSRS("iCustomerID")=iId
oSRS("dTS") = now()
oSRS.Update
oSRS.Close
Set oSRS = Nothing
on error goto 0
end sub
public function currentPopup
'on error resume next
set currentPopup=new cls_popup
dim forcePP
forcePP=convertGetal(left(request.querystring("forcePP"),9))
if forcePP<>0 then
currentPopup.pick(forcePP)
exit function
end if
if isLeeg(session("PopupLoaded")) then
dim rs
set rs=db.execute("select iId from tblPopup where iCustomerID="& cId & " and bEnabled=" & getSQLBoolean(true))
while not rs.eof
currentPopup.pick(rs(0))
if not currentPopup.bOnline then
currentPopup.iId=null
else
currentPopup.iShows=convertGetal(currentPopup.iShows)+1
currentPopup.save()
set cPopup=currentPopup
end if
rs.movenext
wend
set rs=nothing
end if
'on error goto 0
end function
public property get sGetPopupMode
if convertGetal(sPopupViewmode)<>0 then
sGetPopupMode=sPopupViewmode
else
sGetPopupMode="1"
end if
end property
public property get supportZipper
if convertStr(application("supportZipper"))="" then
application("supportZipper")="false"
on error resume next
Err.Clear
dim zippera
Set zippera = Server.CreateObject("aspZip.EasyZIP")
if err.number <>0 then
if instr(sUrl,".")<>0  then
application("supportZipper")="true"
end if
else
application("supportZipper")="true"
end if
Set zippera=nothing
on error goto 0
end if
supportZipper=convertBool(application("supportZipper"))
end property
public property get getPeelDIV
if not bPeelEnabled then exit property
getPeelDIV="<div style=""position:absolute;z-index:2000"" id=""pageflip""><a "
if bPeelOINW then getPeelDIV=getPeelDIV& " target=""_blank"" "
dim psPeelURL
psPeelURL=sPeelURL
psPeelURL=replace(psPeelURL,"&amp;","&",1,-1,1)
psPeelURL=replace(psPeelURL,"&","&amp;",1,-1,1)
getPeelDIV=getPeelDIV& " href=""" & psPeelURL & """><img  style=""border-style:none;border:0px solid #FFF"" src=""" & C_DIRECTORY_QUICKERSITE & "/fixedImages/peels/pageflip" & customer.sPeelFlipColor & ".png"" alt="""" /></a>	<div class=""msg_block""></div></div>"
end property
public sub clearPageCache
on error resume next
if bUseCachingForPages then
dim rs
set rs=db.execute("update tblPage set sPageCache=null where iCustomerID=" & convertGetal(cId))
set rs=nothing
end if
on error goto 0
end sub
public function sCookieJS
if bCookieWarning then
if isleeg(sCWLocation) then
sCWLocation="top"
end if
sCWText=treatconstants(sCWText,true)
sCWAccept=treatconstants(sCWAccept,true)
if not isLeeg(sCWLinkColor) then
sCWText=replace(sCWText,"<a href=","<a style=""text-decoration:underline;color:" & sCWLinkColor & """ href=",1,-1,1)
sCWAccept=replace(sCWAccept,"<a href=","<a style=""text-decoration:underline;color:" & sCWLinkColor & """ href=",1,-1,1)
sCWText=replace(sCWText,"<a class=","<a style=""text-decoration:underline;color:" & sCWLinkColor & """ class=",1,-1,1)
sCWAccept=replace(sCWAccept,"<a class=","<a style=""text-decoration:underline;color:" & sCWLinkColor & """ class=",1,-1,1)
end if
sCookieJS=vbcrlf & "<script type=""text/javascript"" src=""" & C_DIRECTORY_QUICKERSITE & "/js/CookieDirective15.js""></script>" & vbcrlf
sCookieJS=sCookieJS&"<script type=""text/javascript"">" & vbcrlf & "bottomortop='" & sCWLocation & "';" & vbcrlf
sCookieJS=sCookieJS&"sQSVD='" & C_DIRECTORY_QUICKERSITE & "';" & vbcrlf & "</script>" & vbcrlf
if request.querystring("QSbcwp")="1" then
sCookieJS=sCookieJS&"<script type=""text/javascript"">try {cdHandler('" & sCWLocation & "','privacy.html')} catch(error) {};</script>"& vbcrlf
else
sCookieJS=sCookieJS&"<script type=""text/javascript"">try {cookiesDirective('" & sCWLocation & "'," & sCWNumber & ",'privacy.html')} catch (error) {};</script>"& vbcrlf
end if
sCookieJS=sCookieJS&"<script type=""text/javascript"">try {" & vbcrlf
if not isLeeg(sCWText) then
sCookieJS=sCookieJS&"document.getElementById('sCWText').innerHTML=unescape('"  & escape(sCWText) &"');" & vbcrlf
else
sCookieJS=sCookieJS&"document.getElementById('sCWText').style.display='none';" & vbcrlf
end if
sCookieJS=sCookieJS&"document.getElementById('sCWError').innerHTML=unescape('"  & escape(sCWError) &"');" & vbcrlf
sCookieJS=sCookieJS&"document.getElementById('sCWAccept').innerHTML=unescape('"  & escape(sCWAccept) &"');" & vbcrlf
if not isLeeg(sCWContinue) then
sCookieJS=sCookieJS&"document.getElementById('epdsubmit').value=unescape('"  & escape(sCWContinue) &"');" & vbcrlf
end if
if not isLeeg(sCWButtonClass) then
sCookieJS=sCookieJS&"document.getElementById('epdsubmit').className=unescape('"  & escape(sCWButtonClass) &"');" & vbcrlf
end if
if not isLeeg(sCWTextColor) then
sCookieJS=sCookieJS&"document.getElementById('cookiesdirective').style.color='" & sCWTextColor & "';" & vbcrlf
end if
if not isLeeg(sCWBackgroundColor) then
sCookieJS=sCookieJS&"document.getElementById('cookiesdirective').style.background='" & sCWBackgroundColor & "';" & vbcrlf
end if
if convertBool(customer.bCWUseAsNormalPP) then
sCookieJS=sCookieJS&"document.getElementById('bCWUseAsNormalPP').style.display='none';" & vbcrlf
end if
sCookieJS=sCookieJS&"} catch (error) {};</script>"
else
sCookieJS=""
end if
end function

function getArrowJS
	if not isLeeg(sArrowUP) then
		getArrowJS="<a href=""#"" class=""back-to-top""><img alt=""backtotop"" style=""box-shadow:none !important;border-style:none !important"" src=""" & C_DIRECTORY_QUICKERSITE  & "/fixedImages/arrows/" & sArrowUP & """ /></a>"
		getArrowJS=getArrowJS & "<script type=""text/javascript"">jQuery(document).ready(function() {"
		getArrowJS=getArrowJS & "var offset = 220;var duration = 500;jQuery(window).scroll(function() {"
		getArrowJS=getArrowJS & "if (jQuery(this).scrollTop() > offset) {jQuery('.back-to-top').fadeIn(duration);"
		getArrowJS=getArrowJS & "} else {jQuery('.back-to-top').fadeOut(duration);"
		getArrowJS=getArrowJS & "}});jQuery('.back-to-top').click(function(event) {"
		getArrowJS=getArrowJS & "event.preventDefault();jQuery('html, body').animate({scrollTop: 0}, duration);return false;})});</script>"
	end if
end function

function getArrowCSS

	if not isLeeg(sArrowUP) then
		getArrowCSS="<style type=""text/css"">"
		getArrowCSS=getArrowCSS& "div#page {max-width: 900px;margin-left: auto;margin-right: auto;padding: 20px;} "
		getArrowCSS=getArrowCSS& ".back-to-top {position: fixed;bottom: 2em;right: 0px;text-decoration: none;color: #000000;font-size: 12px;padding: 1em;display: none;} "
		getArrowCSS=getArrowCSS& ".back-to-top:hover {}"
		getArrowCSS=getArrowCSS& "</style>"
	end if

end function


end class%>
