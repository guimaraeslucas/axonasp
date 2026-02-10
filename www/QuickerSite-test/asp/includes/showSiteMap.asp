
<%function showSiteMap
set showSiteMap=new cls_menu
showSiteMap.cssType=2
'caching the left menu
if isEmpty(application(QS_CMS_cacheSitemap)) or application(QS_CMS_cacheSitemap)="" then application(QS_CMS_cacheSitemap)=showSiteMap.getMenu(null)
siteMapContent=application(QS_CMS_cacheSitemap)
if convertBool(customer.intranetUse) then
showSiteMap.menu=""
siteMapContent=siteMapContent&showSiteMap.getIntranetMenu(null)
end if
showSiteMap=siteMapContent
end function%>
