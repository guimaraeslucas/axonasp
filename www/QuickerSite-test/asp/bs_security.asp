
<%dim secondAdmin
set secondAdmin=customer.secondAdmin
'pagina ophalen
dim bs_page
bs_page=Request("bs_page")
if not convertBool(Session(cId & "isAUTHENTICATED")) and not convertBool(Session(cId & "isAUTHENTICATEDSecondAdmin")) then
If QS_enableCookieMode then
if not logon.logon(request.cookies.item(cId&"hfsdsiiqqssdfjf")) then Response.Redirect ("noaccess.htm")
else
Response.Redirect ("noaccess.htm")
end if
end if
if not isLeeg(bs_page) then
'rechtstreeks doorsturen naar betrokken pagina.
Response.Redirect (bs_page)
elseif logon.currentPW<>secondAdmin.sPassword then
secondAdmin.bSetupGeneral	= true
secondAdmin.bSetupPageElements	= true
secondAdmin.bStats	= true
secondAdmin.bTemplates	= true
secondAdmin.bHomeConstants	= true
if customer.bApplication then
secondAdmin.bHomeVBScript	= true
else
secondAdmin.bHomeVBScript	= false
end if
secondAdmin.bFiles	= true
secondAdmin.bForms	= true
secondAdmin.bIntranet	= true
secondAdmin.bIntranetSetup	= true
secondAdmin.bIntranetContacts	= true
secondAdmin.bIntranetMail	= true
secondAdmin.bCatalog	= true
secondAdmin.bFeed	= true
secondAdmin.bTheme	= true
secondAdmin.bFormExport	= true
secondAdmin.bApplicationpath	= true
secondAdmin.bGallery	= true
secondAdmin.bPagesAdd	= true
secondAdmin.bPagesPW	= true
secondAdmin.bPagesMove	= true
secondAdmin.bPagesDelete	= true
secondAdmin.bPageUrlRSS	= true
secondAdmin.bPageAdditionalHeader	= true
secondAdmin.bPageDescription	= true
secondAdmin.bPageKeywords	= true
secondAdmin.bPageTitleTag	= true
secondAdmin.bPageRefresh	= true
secondAdmin.bPageTemplate	= true
secondAdmin.bPageTheme	= true
secondAdmin.bPageForm	= true
secondAdmin.bPageCatalog	= true
secondAdmin.bPageFeed	= true
secondAdmin.bPageSetHomepage	= true
secondAdmin.bPagePublish	= true
secondAdmin.bPageOrder	= true
secondAdmin.bPageTitle	= true
secondAdmin.bPageBody	= true
secondAdmin.bPageUFL	= true
secondAdmin.bCalendar	= true
secondAdmin.bPoll	= true
secondAdmin.bGuestbook	= true
secondAdmin.bPopup	= true
secondAdmin.bAvailabilityCal	= true
if customer.bShoppingCart then
secondAdmin.bShoppingCart	= true
else
secondAdmin.bShoppingCart	= false
end if
secondAdmin.bCookieWarning	= true
secondAdmin.bCustom404=true

if customer.bEnableNewsletters then	secondAdmin.bNewsletter = true
end if
if sha256(QS_defaultPW)=customer.adminPassword and blockDefaultPW then 
if not devVersion() then Response.Redirect ("bs_initialsetup.asp")
end if
customer.clearPageCache()%>
