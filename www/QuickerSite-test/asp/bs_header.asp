
<%

function buildMenu (byref arrMenu, btn, iCols)

	dim f

	buildMenu="<table width=""500"" align=center border=0>"		
	buildMenu=buildMenu&"<tr>"
	
	for f=lbound(arrMenu,2) to ubound(arrMenu,2)
		if not isLeeg(arrMenu(0,f)) then
			buildMenu=buildMenu&"<td align=""center"" style=""width:"&round((95/iCols),1)&"%"">"
			
			buildMenu=buildMenu & "<a style=""text-decoration: none;"" "
			buildMenu=buildMenu & "target=" & """" & arrMenu(3,f) & """" & " href=" & """" & arrMenu(1,f) & """" &">"

			'buildMenu=buildMenu & lcase(arrMenu(0,f)) & "<br>"			
			
			buildMenu=buildMenu & "<span class=""material-symbols-outlined"" style=""font-size:40px"">"
			
			
			
			'hier het icoon toevoegen
			select case lcase(arrMenu(0,f))
				
				case "btn_setupbis","btn_setup","btn_setupi" : buildMenu=buildMenu & "settings"
				case "btn_statsbis","btn_stats" : buildMenu=buildMenu & "bar_chart"
				case "btn_securitybis","btn_security" : buildMenu=buildMenu & "security"
				case "btn_pageelementsbis","btn_pageelements" : buildMenu=buildMenu & "widgets"
				case "btn_templatebis","btn_template" : buildMenu=buildMenu & "view_quilt"
				case "btn_contacts" : buildMenu=buildMenu & "group"
				case "btn_checki" : buildMenu=buildMenu & "preview"
				case "btn_sentmessages" : buildMenu=buildMenu & "mail"
				case "btn_intraneti" : buildMenu=buildMenu & "vpn_lock"
				case "btn_theme" : buildMenu=buildMenu & "forum"
				
				
				
				
				
			end select	
			
			
			buildMenu=buildMenu & "</span><br><b>"& arrMenu(2,f) &"</b></a>"	
			
			buildMenu=buildMenu&"</td>"
			
		end if
	next
	
	buildMenu=buildMenu&"</tr></table><br />"

end function

function getBOHeader(btn)
'obsolete
end function
function getMenuArray
dim arrMenu (4,18)
arrMenu (0,0)=btn_Help
arrMenu (1,0)=MYQS_urlSupport
arrMenu (2,0)=l("help")
arrMenu (3,0)="_blank"
arrMenu (4,0)=null
arrMenu (0,1)=btn_Setup
arrMenu (2,1)=l("setup")
arrMenu (3,1)="_self"
arrMenu (4,1)="bosetup"
if secondAdmin.bSetupGeneral then
arrMenu (1,1)="bs_admin.asp"
elseif secondAdmin.bSetupPageElements then
arrMenu (1,1)="bs_pageelements.asp"
elseif secondAdmin.bStats then
arrMenu (1,1)="bs_stats.asp"
elseif secondAdmin.bTemplates then
arrMenu (1,1)="bs_templateList.asp"
else
arrMenu (1,1)="bs_applyTotalPWSA.asp"
end if
arrMenu (0,2)=btn_Home
arrMenu (1,2)="bs_default.asp"
arrMenu (2,2)=l("pagelist")
arrMenu (3,2)="_self"
arrMenu (4,2)="home"
if secondAdmin.bFiles then
arrMenu (0,3)=btn_Folder
arrMenu (1,3)="bs_assetmanager.asp"
arrMenu (2,3)=l("files")
arrMenu (3,3)="_self"
arrMenu (4,3)=null
end if
if secondAdmin.bForms then
arrMenu (0,4)=btn_Forms
arrMenu (1,4)="bs_FormList.asp"
arrMenu (2,4)=l("forms")
arrMenu (3,4)="_self"
arrMenu (4,4)="modules"
end if
if secondAdmin.bIntranet then
arrMenu (0,5)=btn_Intranet
arrMenu (1,5)="bs_Intranet.asp"
arrMenu (2,5)=l("intranet")
arrMenu (3,5)="_self"
arrMenu (4,5)="intranet"
end if
if secondAdmin.bCatalog then
arrMenu (0,12)=btn_Catalog
arrMenu (1,12)="bs_catalogList.asp"
arrMenu (2,12)=l("catalog")
arrMenu (3,12)="_self"
arrMenu (4,12)="modules"
end if
if secondAdmin.bFeed then
arrMenu (0,7)=btn_feed
arrMenu (1,7)="bs_feedList.asp"
arrMenu (2,7)=l("feed")
arrMenu (3,7)="_self"
arrMenu (4,7)="modules"
end if
if secondAdmin.bGallery and QS_ASPX then
arrMenu (0,8)=btn_gallery
arrMenu (1,8)="bs_galleryList.asp"
arrMenu (2,8)=l("gallery")
arrMenu (3,8)="_self"
arrMenu (4,8)="modules"
end if
if secondAdmin.bPoll  then
arrMenu (0,9)=btn_poll
arrMenu (1,9)="bs_pollList.asp"
arrMenu (2,9)="Poll"
arrMenu (3,9)="_self"
arrMenu (4,9)="modules"
end if
if secondAdmin.bGuestbook  then
arrMenu (0,10)=btn_gb
arrMenu (1,10)="bs_gbList.asp"
arrMenu (2,10)=l("guestbook")
arrMenu (3,10)="_self"
arrMenu (4,10)="modules"
end if
if secondAdmin.bPopup  then
arrMenu (0,11)=btn_popup
arrMenu (1,11)="bs_popupList.asp"
arrMenu (2,11)="Popups"
arrMenu (3,11)="_self"
arrMenu (4,11)="modules"
end if
if secondAdmin.bNewsletter  then
arrMenu (0,6)=btn_newsletter
arrMenu (1,6)="bs_newsletterList.asp"
arrMenu (2,6)="Newsletter"
arrMenu (3,6)="_self"
arrMenu (4,6)="modules"
end if
'if secondAdmin.bAvailabilityCal then
'	arrMenu (0,13)=btn_peel
'	arrMenu (1,13)="bs_peellist.asp"
'	arrMenu (2,13)="Peels"
'	arrMenu (3,13)="_self"
'	arrMenu (4,13)="modules"
'end if
if secondAdmin.bAvailabilityCal then
arrMenu (0,14)=btn_ac
arrMenu (1,14)="bs_ac.asp"
arrMenu (2,14)="Availability Calendars"
arrMenu (3,14)="_self"
arrMenu (4,14)="modules"
end if
if secondAdmin.bShoppingCart then
'arrMenu (0,18)=btn_qsc
'arrMenu (1,18)="bs_shoppingcart.asp"
'arrMenu (2,18)="Shopping Cart"
'arrMenu (3,18)="_self"
'arrMenu (4,18)="modules"
end if
arrMenu (0,15)=btn_Check
arrMenu (1,15)="../default.asp?iId=" & encrypt(getHomepage.iId) & "&amp;previewTemplate=false"
arrMenu (2,15)=l("check")
arrMenu (3,15)="_blank"
arrMenu (4,15)=null
arrMenu (0,16)=btn_Exit
arrMenu (1,16)="bs_logoff.asp"
arrMenu (2,16)=l("exit")
arrMenu (3,16)="_self"
arrMenu (4,16)=null
getMenuArray=arrMenu
end function
function getBSArtMenu()
getBSArtMenu="<li><a href=""#""><span class=""l""></span><span class=""r""></span><span class=""t"">" & quotrep(l("login")) & "</span></a></li>"
if bUseArtLoginTemplate then exit function
if convertBool(Session(cId & "isAUTHENTICATEDasADMIN")) then 
if instr(request.servervariables("script_name"),"/ad_")<>0 then
getBSArtMenu="<li><a href=""ad_default.asp""><span class=""l""></span><span class=""r""></span><span class=""t"">Home</span></a></li>"
getBSArtMenu=getBSArtMenu&"<li><a href=""ad_customer.asp""><span class=""l""></span><span class=""r""></span><span class=""t"">New site</span></a></li>"
'getBSArtMenu=getBSArtMenu&"<li><a href=""ad_iis.asp""><span class=""l""></span><span class=""r""></span><span class=""t"">IIS Admin</span></a></li>"
getBSArtMenu=getBSArtMenu&"<li><a href=""ad_clientList.asp""><span class=""l""></span><span class=""r""></span><span class=""t"">Customers</span></a></li>"
getBSArtMenu=getBSArtMenu&"<li><a class=""bPopupFullWidthNoReload"" href=""ad_labels.asp""><span class=""l""></span><span class=""r""></span><span class=""t"">Labels</span></a><ul><li><a class=""bPopupFullWidthNoReload"" href=""ad_label.asp"">New label</a></li></ul></li>"
getBSArtMenu=getBSArtMenu&"<li><a href=""#""><span class=""l""></span><span class=""r""></span><span class=""t"">Monitor</span></a><ul><li><a href=""ad_pageList.asp"">Pages</a></li><li><a href=""ad_refererList.asp"">Referrers</a></li><li><a href=""ad_gbentries.asp"">Guestbooks</a></li><li><a href=""ad_formentries.asp"">Forms</a></li><li><a href=""ad_monitor.asp"">Hits</a></li></ul></li>"
getBSArtMenu=getBSArtMenu&"<li><a href=""ad_logoff.asp""><span class=""l""></span><span class=""r""></span><span class=""t"">Exit</span></a></li>"
exit function
end if
end if
if not convertBool(Session(cId & "isAUTHENTICATED")) and not convertBool(Session(cId & "isAUTHENTICATEDSecondAdmin")) then
If QS_enableCookieMode then
if not logon.logon(request.cookies.item(cId&"hfsdsiiqqssdfjf")) then 
exit function
end if
else
exit function
end if
end if
getBSArtMenu=""
'if isLeeg(session(cId & "getMenuArray")) then
dim arrMenu,bModulesLoaded,modulesSubmenu,arrIntranetMenu,arrBOsetup,TemplateMenuArr,PEMenuArr,NLMenuArr,HomeMenuArr,FilesMenuArr
arrMenu=getMenuArray
arrIntranetMenu=getIntranetMenuArray
arrBOsetup=getBOSetupMenuArr
TemplateMenuArr=getTemplateMenuArr
NLMenuArr=getNLMenuArr
PEMenuArr=getPEMenuArr
HomeMenuArr=getHomeMenuArr
bModulesLoaded=false
modulesSubmenu=""
dim arrM,arrIm,arrBOS,arrT,arrP,arrNL,arrH,arrF
for arrM=lbound(arrMenu,2) to ubound(arrMenu,2)
if isNull(arrMenu(4,arrM)) then
getBSArtMenu=getBSArtMenu& "<li><a target=""" & quotrep(arrMenu(3,arrM)) & """ href=""" & quotrep(arrMenu(1,arrM)) & """><span class=""l""></span><span class=""r""></span><span class=""t"">" & quotrep(arrMenu(2,arrM)) & "</span></a></li>"
elseif convertStr(arrMenu(4,arrM))="intranet" then
getBSArtMenu=getBSArtMenu& "<li><a target=""" & quotrep(arrMenu(3,arrM)) & """ href=""" & quotrep(arrMenu(1,arrM)) & """><span class=""l""></span><span class=""r""></span><span class=""t"">" & quotrep(arrMenu(2,arrM)) & "</span></a><ul>"
for arrIm=lbound(arrIntranetMenu,2) to ubound(arrIntranetMenu,2)
if not isLeeg(arrIntranetMenu(2,arrIm)) then
getBSArtMenu=getBSArtMenu& "<li><a target=""" & quotrep(arrIntranetMenu(3,arrIm)) & """ href=""" & quotrep(arrIntranetMenu(1,arrIm)) & """>" & quotrep(arrIntranetMenu(2,arrIm)) & "</a></li>"
end if
next
getBSArtMenu=getBSArtMenu& "</ul></li>"
elseif convertStr(arrMenu(4,arrM))="home" then
getBSArtMenu=getBSArtMenu& "<li><a target=""" & quotrep(arrMenu(3,arrM)) & """ href=""" & quotrep(arrMenu(1,arrM)) & """><span class=""l""></span><span class=""r""></span><span class=""t"">" & quotrep(arrMenu(2,arrM)) & "</span></a><ul>"
for arrH=lbound(HomeMenuArr,2) to ubound(HomeMenuArr,2)
if not isLeeg(HomeMenuArr(2,arrH)) then
getBSArtMenu=getBSArtMenu& "<li><a target=""" & quotrep(HomeMenuArr(3,arrH)) & """ href=""" & quotrep(HomeMenuArr(1,arrH)) & """>" & quotrep(HomeMenuArr(2,arrH)) & "</a></li>"
end if
next
getBSArtMenu=getBSArtMenu& "</ul></li>"
elseif convertStr(arrMenu(4,arrM))="bosetup" then
getBSArtMenu=getBSArtMenu& "<li><a target=""" & quotrep(arrMenu(3,arrM)) & """ href=""" & quotrep(arrMenu(1,arrM)) & """><span class=""l""></span><span class=""r""></span><span class=""t"">" & quotrep(arrMenu(2,arrM)) & "</span></a><ul>"
for arrBOS=lbound(arrBOsetup,2) to ubound(arrBOsetup,2)
if not isLeeg(arrBOsetup(2,arrBOS)) then
getBSArtMenu=getBSArtMenu& "<li><a target=""" & quotrep(arrBOsetup(3,arrBOS)) & """ href=""" & quotrep(arrBOsetup(1,arrBOS)) & """>" & quotrep(arrBOsetup(2,arrBOS)) & "</a><ul>"
if arrBOsetup(0,arrBOS)=btn_template then
for arrT=lbound(TemplateMenuArr,2) to ubound (TemplateMenuArr,2)
if not isLeeg(TemplateMenuArr(2,arrT)) then
getBSArtMenu=getBSArtMenu& "<li><a target=""" & quotrep(TemplateMenuArr(3,arrT)) & """ href=""" & quotrep(TemplateMenuArr(1,arrT)) & """>" & quotrep(TemplateMenuArr(2,arrT)) & "</a></li>"
end if
next
end if
if arrBOsetup(0,arrBOS)=btn_pageelements then
for arrP=lbound(PEMenuArr,2) to ubound (PEMenuArr,2)
if not isLeeg(PEMenuArr(2,arrP)) then
getBSArtMenu=getBSArtMenu& "<li><a target=""" & quotrep(PEMenuArr(3,arrP)) & """ href=""" & quotrep(PEMenuArr(1,arrP)) & """>" & quotrep(PEMenuArr(2,arrP)) & "</a></li>"
end if
next
end if
getBSArtMenu=getBSArtMenu& "</ul></li>"
end if
next
getBSArtMenu=getBSArtMenu& "</ul></li>"
elseif convertStr(arrMenu(4,arrM))="modules" then
if not bModulesLoaded then
getBSArtMenu=getBSArtMenu& "++labelmodules++"
bModulesLoaded=true
end if
modulesSubmenu=modulesSubmenu& "<li><a target=""" & quotrep(arrMenu(3,arrM)) & """ href=""" & quotrep(arrMenu(1,arrM)) & """>" & quotrep(arrMenu(2,arrM)) & "</a><ul>"
if arrMenu(0,arrM)=btn_newsletter then
for arrNL=lbound(NLMenuARR,2) to ubound (NLMenuARR,2)
if not isLeeg(NLMenuARR(2,arrNL)) then
modulesSubmenu=modulesSubmenu& "<li><a target=""" & quotrep(NLMenuARR(3,arrNL)) & """ href=""" & quotrep(NLMenuARR(1,arrNL)) & """>" & quotrep(NLMenuARR(2,arrNL)) & "</a></li>"
end if
next
end if
modulesSubmenu=modulesSubmenu& "</ul></li>"
end if
next
if not isLeeg(modulesSubmenu) then
getBSArtMenu=replace(getBSArtMenu,"++labelmodules++","<li><a href=""#""><span class=""l""></span><span class=""r""></span><span class=""t"">Modules</span></a><ul>" & modulesSubmenu & "</ul></li>",1,-1,1)
end if
getBSArtMenu=replace(getBSArtMenu,"<ul></ul>","",1,-1,1)
'getBSArtMenu=replace(getBSArtMenu,"target=""_blank"""," class=""bPopupFullWidthNoReload"" ",1,-1,1)
'getBSArtMenu=replace(getBSArtMenu,"target=""_blankSmall"""," class=""QSPP"" ",1,-1,1)
session(cId & "getMenuArray")=getBSArtMenu
'end if
getBSArtMenu=session(cId & "getMenuArray")
end function
function getIntranetMenuArray
dim arrMenu (3,5)
if secondAdmin.bIntranet then
arrMenu (0,0)=btn_IntranetI
arrMenu (1,0)="bs_Intranet.asp"
arrMenu (2,0)=l("intranet")
arrMenu (3,0)="_self"
end if
if secondAdmin.bIntranetSetup then
arrMenu (0,1)=btn_SetupI
arrMenu (1,1)="bs_adminIntranet.asp"
arrMenu (2,1)=l("setup")
arrMenu (3,1)="_self"
end if
if secondAdmin.bTheme then
arrMenu (0,2)=btn_theme
arrMenu (1,2)="bs_themeslist.asp"
arrMenu (2,2)=l("themes")
arrMenu (3,2)="_self"
end if
if secondAdmin.bIntranetContacts then
arrMenu (0,3)=btn_contacts
arrMenu (1,3)="bs_contactHome.asp"
arrMenu (2,3)=l("contacts")
arrMenu (3,3)="_self"
end if
if secondAdmin.bIntranetMail then
arrMenu (0,4)=btn_sentMessages
arrMenu (1,4)="bs_mailHistory.asp"
arrMenu (2,4)=l("messages")
arrMenu (3,4)="_self"
end if
arrMenu (0,5)=btn_CheckI
arrMenu (1,5)="default.asp?iId="& encrypt(getIntranetHomePage.iId)
arrMenu (2,5)=l("check")
arrMenu (3,5)="_blank"
getIntranetMenuArray=arrMenu
end function
function getBOHeaderIntranet(btn)
Response.write buildMenu(getIntranetMenuArray,btn,6)
end function
function getBOSetupMenu(btn)
Response.write buildMenu(getBOSetupMenuArr,btn,7)
end function
function getBOSetupMenuArr
dim arrMenu (4,16)
if secondAdmin.bSetupGeneral then
arrMenu (0,0)=btn_SetupBis
arrMenu (1,0)="bs_admin.asp"
arrMenu (2,0)=l("general")
arrMenu (3,0)="_self"
end if
if secondAdmin.bSetupPageElements then
arrMenu (0,3)=btn_pageelements
arrMenu (1,3)="bs_pageelements.asp"
arrMenu (2,3)=l("pageelements")
arrMenu (3,3)="_self"
end if
if secondAdmin.bStats then
arrMenu (0,4)=btn_Stats
arrMenu (1,4)="bs_stats.asp"
arrMenu (2,4)=l("stats")
arrMenu (3,4)="_self"
end if
if customer.bScanreferer and secondAdmin.bStats then
arrMenu (0,5)=""
arrMenu (1,5)="bs_referers.asp"
arrMenu (2,5)=l("referringsites")
arrMenu (3,5)="_self"
end if
if secondAdmin.bPagesPW then
arrMenu (0,6)=btn_Security
if logon.currentPW<>customer.secondAdmin.sPassword then
arrMenu (1,6)="bs_applyTotalPW.asp"
else
arrMenu (1,6)="bs_applyTotalPWSA.asp"
end if
arrMenu (2,6)=l("security")
arrMenu (3,6)="_self"
else
arrMenu (0,7)=btn_Security
arrMenu (1,7)="bs_applyTotalPWSA.asp"
arrMenu (2,7)=l("security")
arrMenu (3,7)="_self"
end if
if logon.currentPW<>customer.secondAdmin.sPassword then
arrMenu (0,8)=""
arrMenu (1,8)="bs_secondAdmin.asp"
arrMenu (2,8)=l("secondadmin")
arrMenu (3,8)="_self"
end if
if secondAdmin.bTemplates then
arrMenu (0,10)=""
arrMenu (1,10)="bs_popupMode.asp"
arrMenu (2,10)="Popup-effect"
arrMenu (3,10)="_self"
end if
if secondAdmin.bTemplates then
arrMenu (0,13)=""
arrMenu (1,13)="bs_AccordionSetup.asp"
arrMenu (2,13)="Accordion CSS"
arrMenu (3,13)="_self"
end if
if secondAdmin.bTemplates then
arrMenu (0,9)=btn_template
arrMenu (1,9)="bs_templateList.asp"
arrMenu (2,9)=l("templates")
arrMenu (3,9)="_self"
end if
if secondAdmin.bTemplates then
arrMenu (0,11)=""
arrMenu (1,11)="bs_mobileSetup.asp"
arrMenu (2,11)="Mobile setup"
arrMenu (3,11)="_self"
end if
if secondAdmin.bCookieWarning then
arrMenu (0,12)=""
arrMenu (1,12)="bs_cookiewarning.asp"
arrMenu (2,12)="Cookie warning"
arrMenu (3,12)="_self"
end if
if customer.bUserFriendlyURL then
if secondAdmin.bCustom404 then
arrMenu (0,14)=""
arrMenu (1,14)="bs_404.asp"
arrMenu (2,14)="Custom 404 page"
arrMenu (3,14)="_self"
end if
end if
arrMenu (0,15)=""
arrMenu (1,15)="bs_arrowup.asp"
arrMenu (2,15)="Scroll To Top"
arrMenu (3,15)="_self"

getBOSetupMenuArr=arrMenu
end function
function getTemplateMenuArr
dim arrMenu (3,4)
arrMenu (0,0)=""
arrMenu (1,0)="bs_templateList.asp"
arrMenu (2,0)="List of templates"
arrMenu (3,0)="_self"
arrMenu (0,2)=""
arrMenu (1,2)="bs_templateEdit.asp"
arrMenu (2,2)=l("newtemplate")
arrMenu (3,2)="_self"
if customer.supportZipper then
arrMenu (0,4)=""
arrMenu (1,4)="bs_uploadzip.asp"
arrMenu (2,4)="Upload zipped template"
arrMenu (3,4)="_self"
end if
if bBrowseOnlineTemplates then
arrMenu (0,3)=""
arrMenu (1,3)="bs_templateSearch.asp"
arrMenu (2,3)="Search for templates online"
arrMenu (3,3)="_self"
end if
getTemplateMenuArr=arrMenu
end function
function getPEMenuArr
dim arrMenu (4,5)
arrMenu (0,1)=""
arrMenu (1,1)="bs_editbannermenu.asp"
arrMenu (2,1)=l("banners")
arrMenu (3,1)="_self"
arrMenu (0,2)=""
arrMenu (1,2)="bs_editprops.asp"
arrMenu (2,2)="Default Blocks"
arrMenu (3,2)="_self"
arrMenu (0,3)=""
arrMenu (1,3)="bs_editfooter.asp"
arrMenu (2,3)=l("footer")
arrMenu (3,3)="_self"
arrMenu (0,4)=""
arrMenu (1,4)="bs_favicon.asp"
arrMenu (2,4)=l("favicon")
arrMenu (3,4)="_self"
arrMenu (0,5)=""
arrMenu (1,5)="bs_peel.asp"
arrMenu (2,5)="Peel"
arrMenu (3,5)="_self"
getPEMenuArr=arrMenu
end function
function getNLMenuArr
dim arrMenu (3,10)
arrMenu (0,0)=""
arrMenu (1,0)="bs_NewsletterList.asp"
arrMenu (2,0)="List Newsletters"
arrMenu (3,0)="_self"
arrMenu (0,1)=""
arrMenu (1,1)="bs_NewsletterEdit.asp"
arrMenu (2,1)="Create New Newsletter"
arrMenu (3,1)="_self"
if customer.bCanImportSubscribers then
arrMenu (0,3)=""
arrMenu (1,3)="bs_NewsletterMailingHistory.asp"
arrMenu (2,3)="Mailing History"
arrMenu (3,3)="_self"
arrMenu (0,5)=""
arrMenu (1,5)="bs_NewsletterSubscribers.asp"
arrMenu (2,5)="Subscribers"
arrMenu (3,5)="_self"
arrMenu (0,7)=""
arrMenu (1,7)="bs_NewsletterImport.asp"
arrMenu (2,7)="Import Subscribers"
arrMenu (3,7)="_self"
arrMenu (0,9)=""
arrMenu (1,9)="bs_NewsletterCategoryList.asp"
arrMenu (2,9)="Email lists"
arrMenu (3,9)="_self"
else
arrMenu (0,9)=""
arrMenu (1,9)="bs_NewsletterCategoryEdit.asp"
arrMenu (2,9)="Create new Email list"
arrMenu (3,9)="_self"
end if
getNLMenuArr=arrMenu
end function
function getHomeMenuArr
dim arrMenu (3,4)
arrMenu (0,0)=""
arrMenu (1,0)="bs_default.asp"
arrMenu (2,0)=l("pagelist")
arrMenu (3,0)="_self"
if secondAdmin.bPagesAdd then
arrMenu (0,1)=""
arrMenu (1,1)="bs_setupPage.asp"
arrMenu (2,1)=l("newpage")
arrMenu (3,1)="_self"
end if
if secondAdmin.bHomeConstants then
arrMenu (0,2)=""
arrMenu (1,2)="bs_ConstantList.asp"
arrMenu (2,2)=l("constants")
arrMenu (3,2)="_self"
end if
if secondAdmin.bHomeVBScript then
arrMenu (0,3)=""
arrMenu (1,3)="bs_scriptlist.asp"
arrMenu (2,3)="ASP/VBScripts"
arrMenu (3,3)="_self"
end if
if logon.currentPW<>customer.secondAdmin.sPassword then
arrMenu (0,4)=""
arrMenu (1,4)="bs_search.asp"
arrMenu (2,4)=l("search")
arrMenu (3,4)="_self"
end if
getHomeMenuArr=arrMenu
end function%>
