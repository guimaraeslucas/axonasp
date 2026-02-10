
<%public function showBreadCrumbs
if not logon.authenticatedIntranet and selectedPage.bIntranet then exit function
if isNumeriek(selectedPage.iId) then
set hereList=new cls_menu
hereList.bOnline=true
smTable=replace(hereList.showParents(selectedPage.iId,true),selectedPage.sTitle& " &gt; ","")
set hereList=nothing
end if
set subPages=selectedPage.subPages(true)
continue=false
for each subPage in subPages
if not subPages(subPage).bContainerPage then
continue=true
end if
next 
if subPages.count>0 and continue then
for each subPage in subPages
if not subPages(subPage).bContainerPage then
gotoMenu=subPages(subPage).getParentLink &" / "& gotoMenu
end if
next
if not isLeeg(gotoMenu) then
smTable=smTable & " &gt; " & left(gotoMenu,len(gotoMenu)-3)
end if
end if
if selectedPage.parentPage.bContainerPage then
smTable=smTable &" / "
set onlineSiblings=selectedPage.onlineSiblings
if onlineSiblings.count>0 then
for each subPage in onlineSiblings
if not onlineSiblings(subPage).bContainerPage then
siblingmenu=onlineSiblings(subPage).getLink(false) &" / "& siblingmenu
end if
next
if not isLeeg(siblingmenu) then
smTable=smTable & left(siblingmenu,len(siblingmenu)-3)
end if
end if
end if
if not isLeeg(smTable) and isLeeg(selectedPage.iTemplateID) and isLeeg(customer.defaultTemplate) then
smTable="<div id=""QS_breadcrumbs"">"&smTable&"</div>"
end if
showBreadCrumbs=smTable
showBreadCrumbs=replace(showBreadCrumbs," id='"," id='QQ",1,-1,1)
end function%>
