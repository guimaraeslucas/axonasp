
<%const save_listitem="SaveListItem"
const save_listpage="SaveListPage"
const clogin="Binnen"
const cloginIntranet="BinnenI"
const cForgotPW="forgotPW"
const cLogOff="logoff"
const cSaveAdminPW="cSaveAdminPW"
const cRegister="register"
const cProfile="profile"
const cWelcome="welcome"
const cPostTopic="posttopic"
const cPostReply="postreply"
const cModReply="modreply"
const cRemoveReply="removereply"
const cSubscribeToTopic="subscribetopic"
const cUnSubscribeToTopic="unsubscribetopic"
const cSubscribeToTheme="subscribetheme"
const cUnSubscribeToTheme="unsubscribetheme"
const cSearchTheme="searchtheme"
const clSubs="lsubs"
const cdSub="cdsub"
const cValidateReply="cValidateReply"
const cValidateTopic="cValidateTopic"
'action Types
const SB_accessAllow="SB_accessAllow"
const SB_accessDeny="SB_accessDeny"
const SB_massMail="SB_massMail"
const SB_sendMail="SB_sendMail"
'menu BO
const btn_Help	= "btn_Help"
const btn_Home	= "btn_Home"
const btn_Setup	= "btn_Setup"
const btn_SetupI	= "btn_SetupI"
const btn_SetupBis	= "btn_SetupBis"
const btn_Banner	= "btn_Banner"
const btn_Logo	= "btn_Logo"
const btn_Favicon	= "btn_Favicon"
const btn_Stats	= "btn_Stats"
const btn_Security	= "btn_Security"
const btn_Intranet	= "btn_Intranet"
const btn_IntranetI	= "btn_IntranetI"
const btn_Catalog	= "btn_Catalog"
const btn_Check	= "btn_Check"
const btn_CheckI	= "btn_CheckI"
const btn_Exit	= "btn_Exit"
const btn_Contacts	= "btn_Contacts"
const btn_sentMessages	= "btn_sentmessages"
const btn_Forms	= "btn_Forms"
const btn_Design	= "btn_Design"
const btn_feed	= "btn_feed"
const btn_template	= "btn_template"
const btn_folder	= "btn_folder"
const btn_pageelements	= "btn_pageelements"
const btn_theme	= "btn_theme"
const btn_gallery	= "btn_gallery"
const btn_calendar	= "btn_calendar"
const btn_poll	= "btn_poll"
const btn_gb	= "btn_gb"
const btn_popup	= "btn_popup"
const btn_peel	= "btn_peel"
const btn_newsletter	= "btn_newsletter"
const btn_ac	= "btn_ac"
const btn_qsc	= "btn_qsc"
const btn_blocks	= "btn_blocks"
const pl_Vertical = "V"
const pl_Horizontal = "H"
'aligment
const QS_centerAlign	= "center"
const QS_rightAlign	= "right"
const QS_leftAlign	= "left"
const QS_topAlign	= "top"
const QS_topRight	= "topRight"
const QS_nomenu	= "nomenu"
const QS_QleftAright	= "QS_QleftAright"
const QS_QtopAbottom	= "QS_QtopAbottom"
const QS_QrightAleft	= "QS_QrightAleft"
'formatting
const QS_textonly	= 0
const QS_html	= 1
const QS_VBScript	= 2
const QS_VBScriptIdentifier = "#!!QS_VBSCRIPT!!#"
'default password
const QS_defaultPW	= "admin"
const QS_feedNoText=10000
'text direction
const QS_rtl="rtl"
const QS_ltr="ltr"
'Caching strings
dim QS_CMS_cacheSitemap, QS_CMS_cacheMenu, QS_CMS_FCK_authCode, QS_CMS_FCK_allowedIP, QS_CMS_FCK_authKey, QS_CMS_C_VIRT_DIR
dim QS_CMS_lb_, QS_CMS_arrfeeds, QS_CMS_feedsreloaded, QS_CMS_themesreloaded, QS_CMS_arrthemes, QS_CMS_cacheBOMenu, QS_textDirection
dim QS_langcode, QS_CMS_cacheFEED, QS_CMS_arrgalleries, QS_CMS_galleriesreloaded, QS_CMS_cacheGallery, QS_CMS_cacheMainMenu,QS_CMS_cacheIntranetMenu
QS_CMS_cacheSitemap	= "QS_CMS_cacheSitemap" & cId
QS_CMS_cacheMenu	= "QS_CMS_cacheMenu" & cId
QS_CMS_FCK_authCode	= "QS_CMS_FCK_authCode" & cId
QS_CMS_FCK_allowedIP	= "QS_CMS_FCK_allowedIP" & cId
QS_CMS_FCK_authKey	= "QS_CMS_FCK_authKey" & cId
QS_CMS_lb_	= "QS_CMS_lb_" & cId
QS_CMS_arrfeeds	= "QS_CMS_arrfeeds" & cId
QS_CMS_feedsreloaded	= "QS_CMS_feedsreloaded" & cId
QS_CMS_themesreloaded	= "QS_CMS_themesreloaded" & cId
QS_CMS_arrthemes	= "QS_CMS_arrthemes" & cId
QS_CMS_cacheBOMenu	= "QS_CMS_cacheBOMenu" & cId 
QS_textDirection	= "QS_textDirection" & cId
QS_langcode	= "QS_langcode" & cId
QS_CMS_cacheFEED	= "QS_CMS_cacheFEED" & cId
QS_CMS_arrgalleries	= "QS_CMS_arrgalleries" & cId
QS_CMS_galleriesreloaded	= "QS_CMS_galleriesreloaded" & cId
QS_CMS_cacheGallery	= "QS_CMS_cacheGallery" & cId
QS_CMS_cacheMainMenu	= "QS_CMS_cacheMainMenu" & cId
QS_CMS_cacheIntranetMenu	= "QS_CMS_cacheIntranetMenu"& cId
'misc constants
dim QS_CMS_hiddenfield
QS_CMS_hiddenfield	= "Hidden field"%>
