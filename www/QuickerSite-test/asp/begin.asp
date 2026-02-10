<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit

on error resume next
dim blockURL : blockURL=Request.ServerVariables("query_string")
if instr(blockURL,"?")>0 then blockURL=left(blockURL,instr(blockURL,"?")-1) : end if
select case trim(lcase(GetFileExtension(blockURL)))
	case "php","log","exe","jsp","jspx","jws","config","py","pyc","pyd","wsgi","xml","svg","cfm","cfml","cgi","pl","pm","phtml","php3","php4","php5","php6","php7","js","bat","cmd","hta","scr","inc","pht","asa","asax","aspx","vbs","vbe","wsf","docm","xlsm","pptm","swf","xap","scf","resx","ico"
	response.Status="404 File Not Found" : response.end
end select
on error goto 0

startTimer=timer()
Response.Buffer				= true
session.Timeout				= 30
server.ScriptTimeout		= 800
const C_QS_VERSION			= "4.4"
const QS_CHARSET			= "utf-8"
const C_FCKEDITOR			= "fckeditor266"
const C_CUTEEDITOR			= "CuteEditor66"
const C_CKEDITOR			= "ckeditor445"
const C_INNOVAEDITOR		= "InnovaStudio35"
editPage					= false
defaultasp					= "default.asp"

QS_ASPX						= true 'use aspx features?

dim QS_EDITOR
Response.CharSet			= QS_CHARSET
Response.ContentType		= "text/html"
Response.CacheControl		= "no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires			= -1
Response.ExpiresAbsolute	= Now()-1

blockDefaultPW		= true
printReplies		= false
pagetoemail			= false
manyContacts		= 500

dim maxPictureSize,resizePictureToPx,pictureResizeSecCode,MYQS_offlineLinkColor,qsscart,iLPDEFAULTOpenerQS, QSCDO_smtpusessl,includeNS
'v27: Resize pictures?
qsscart="shoppingcart"
bBrowseOnlineTemplates=false
includeNS=false
sBrowseOnlineTemplatesUrl="https://templates31.quickersite.com/"
MYQS_offlineLinkColor="#AAAAAA"
iLPDEFAULTOpenerQS=0
QSCDO_smtpusessl=false

dim headerDictionary
set headerDictionary=server.createobject("scripting.dictionary")

%>
<!-- #include file="config/rebrand.asp"-->
<!-- #include file="config/web_config.asp"-->
<!-- #include file="includes/artisteer.asp"-->
<!-- #include file="includes/sha256.asp"-->
<!-- #include file="includes/iso-8859-1.asp"-->
<!-- #include file="includes/fileexplorer.asp"-->
<!-- #include file="includes/constants.asp"-->
<!-- #include file="includes/dateFormatList.asp"-->
<!-- #include file="includes/excelFile.asp"-->
<!-- #include file="includes/messages.asp"-->
<!-- #include file="includes/database.asp"-->
<!-- #include file="includes/logVisit.asp"-->
<!-- #include file="includes/mail_message.asp"-->
<!-- #include file="includes/encryption.asp"-->
<!-- #include file="includes/fixedTypeList.asp"-->
<!-- #include file="includes/formatList.asp"-->
<!-- #include file="includes/alignList.asp"-->
<!-- #include file="includes/menualignList.asp"-->
<!-- #include file="includes/placementList.asp"-->
<!-- #include file="includes/fixedFieldTypeList.asp"-->
<!-- #include file="includes/contactFieldTypeList.asp"-->
<!-- #include file="includes/formFieldTypeList.asp"-->
<!-- #include file="includes/fontTypeList.asp"-->
<!-- #include file="includes/URLTypeList.asp"-->
<!-- #include file="includes/URLTypeShortList.asp"-->
<!-- #include file="includes/logonEdit.asp"-->
<!-- #include file="includes/functions.asp"-->
<!-- #include file="includes/insertConstants.asp"-->
<!-- #include file="includes/getVisitorDetails.asp"-->
<!-- #include file="includes/menu.asp"-->
<!-- #include file="includes/customer.asp"-->
<!-- #include file="includes/customerList.asp"-->
<!-- #include file="includes/freeASPUpload.asp"-->
<!-- #include file="includes/page.asp"-->
<!-- #include file="includes/search.asp"-->
<!-- #include file="includes/languageList.asp"-->
<!-- #include file="includes/language.asp"-->
<!-- #include file="includes/label.asp"-->
<!-- #include file="includes/labelList.asp"-->
<!-- #include file="includes/languageInit.asp"-->
<!-- #include file="includes/orderBYlist.asp"-->
<!-- #include file="includes/catalogOrderBYlist.asp"-->
<!-- #include file="includes/contact.asp"-->
<!-- #include file="includes/contactStatusList.asp"-->
<!-- #include file="includes/contactField.asp"-->
<!-- #include file="includes/contactSearch.asp"-->
<!-- #include file="includes/lowerGreaterThanList.asp"-->
<!-- #include file="includes/mail.asp"-->
<!-- #include file="includes/catalog.asp"-->
<!-- #include file="includes/catalogFileType.asp"-->
<!-- #include file="includes/catalogField.asp"-->
<!-- #include file="includes/catalogItem.asp"-->
<!-- #include file="includes/catalogItemFile.asp"-->
<!-- #include file="includes/groupRefByList.asp"-->
<!-- #include file="includes/itemSearch.asp"-->
<!-- #include file="includes/form.asp"-->
<!-- #include file="includes/formfield.asp"-->
<!-- #include file="includes/submission.asp"-->
<!-- #include file="includes/customerIntranetMessage.asp"-->
<!-- #include file="includes/secondAdmin.asp"-->
<!-- #include file="includes/feed.asp"-->
<!-- #include file="includes/gallery.asp"-->
<!-- #include file="includes/template.asp"-->
<!-- #include file="includes/themeSubscriptionList.asp"-->
<!-- #include file="includes/theme.asp"-->
<!-- #include file="includes/post.asp"-->
<!-- #include file="includes/themeTypeList.asp"-->
<!-- #include file="includes/ticket.asp"-->
<!-- #include file="includes/constant.asp"-->
<!-- #include file="includes/showMenu.asp"-->
<!-- #include file="includes/showSiteMap.asp"-->
<!-- #include file="includes/ckeditor.asp"-->
<!-- #include file="includes/breadcrumbs.asp"-->
<!-- #include file="includes/fullsearch.asp"-->
<!-- #include file="includes/googlemap.asp"-->
<!-- #include file="includes/siteSearchTypeList.asp"-->
<!-- #include file="includes/galleryTypeList.asp"-->
<!-- #include file="includes/gallerySEList.asp"-->
<!-- #include file="includes/zip.asp"-->
<!-- #include file="includes/poll.asp"-->
<!-- #include file="includes/iisobject.asp"-->
<!-- #include file="includes/guestbook.asp"-->
<!-- #include file="includes/guestbookitem.asp"-->
<!-- #include file="includes/popup.asp"-->
<!-- #include file="includes/popupModeList.asp"-->
<!-- #include file="includes/newsletter.asp"-->
<!-- #include file="includes/newsletterMailing.asp"-->
<!-- #include file="includes/newsletterCategory.asp"-->
<!-- #include file="includes/fastpage.asp"-->
<!-- #include file="includes/showicons.asp"-->
<!-- #include file="includes/md5.asp"-->
<!-- #include file="includes/shopMake.asp"-->
<!-- #include file="includes/shopCategory.asp"-->
<!-- #include file="includes/shopProduct.asp"-->
<%

'load customer
dim customer
set customer=new cls_customer
customer.pick(cId)

dim startTimer,SQL2005_SERVER,SQL2005_DB,SQL2005_UID,SQL2005_PWD, blockDefaultPW, printReplies, manyContacts,sNewTemplatesURL,bUseOneIISsite,iForceReload,bAddImageToEmail,sAddImageUrl,sBrowseOnlineTemplatesUrl
dim execBeforePageLoad, execAfterPageLoad, QS_ASPX, QS_DBS, C_DATABASE, C_DATABASE_LABELS,MySQL_PORT,MySQL_SERVER,sNewTemplatesPath,bLoadConstants,sOWBodyBGColor,bUseArtLoginTemplate,sCBExtUrl,saveHiddenValues
dim MySQL_DB,MySQL_UID,MySQL_PWD,MySQL_OPTION, pagetoemail, pagetoemailbody, QS_enableCookieMode, editPage, bThemeEmbed, defaultasp,sMetaTagRefresh,cPopup,bNOPopup,bBrowseOnlineTemplates,i404TemplateID
bLoadConstants=true
bThemeEmbed=false
bUseOneIISsite=false
bAddImageToEmail=false
bUseArtLoginTemplate=false
bNOPopup=false
saveHiddenValues=false
iForceReload=0
QS_EDITOR=4 '4: ckEditor
QS_ASPX=true 'use ASP.NET features - basically the picture thumbnailer
QS_enableCookieMode=true 'set to true in case your host recycles your application every x minutes

'constants
dim QS_CMS_arrconstants, QS_CMS_constantsreloaded
QS_CMS_arrconstants=cstr("QS_CMS_arrconstants"&getDateWithoutSlash&cId)
QS_CMS_constantsreloaded=cstr("QS_CMS_constantsreloaded"&getDateWithoutSlash&cId)

'On or offline?
if C_DEV then
	Response.Write l("QuickerSiteoffline")
	Response.End 
end if

'logonClass
dim logon
set logon=new cls_logonEdit

'Popup
set cPopup=new cls_popup

'load selectedpage
dim selectedPage
set selectedPage=new cls_page
selectedPage.pick(decrypt(left(Request("iId"),40)))
if isNull(selectedPage.iId) then	
	selectedPage.pickByCode(left(request("sCode"),25))
end if

dim cslist, csShortList
set cslist=new cls_contactStatusList

%>
<!-- #include file="includes/javascript.asp"-->
<!-- #include file="includes/css.asp"-->
