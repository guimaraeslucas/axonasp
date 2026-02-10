
<%function showMenu
dim menu
set menu=new cls_menu
showMenu=menu.getMenu(null)
showMenu=showMenu& menu.getIntranetMenu(null)
set menu=nothing
showMenu=treatConstants(showMenu,true)
showMenu=setActiveLink(showMenu)
end function

function getBootstrapMenu(version)

	dim m

	select case version

		case 3
			
			m=showMenu
			m=replace(m,"<ul id='QS_menulist'>","<ul class='nav navbar-nav'>",1,1,1)
			m=replace(m,"</ul></div><div id='menuI'><ul id='QS_intranetmenulist'>","",1,-1,1)
			m=replace(m,"<li><a class='subcontainer'","<li class='dropdown'><a href=# class=dropdown-toggle data-toggle=dropdown role=button aria-haspopup=true aria-expanded=false class='subcontainer'",1)
			m=replace(m,"</a><ul><li>","<span class='caret'></span></a><ul class='dropdown-menu'><li>",1)
			m=JSremoveIDs(m)
			m=setActiveLink(m)
			m=replace(m,"id='qs_activelink'","class=""active""",1,-1,1)
			m=replace(m,"<li><a  class=""active""","<li class=""active""><a ",1,-1,1)
			m=replace(m,"<div >","",1,-1,1)
			m=replace(m,"</div>","",1,-1,1)
			getBootstrapMenu=m

	end select 

end function


function JStemplateVMENU
JStemplateVMENU=showMenu
if instr(JStemplateVMENU,"QS_intranetmenulist")<>0 then
if logon.authenticatedIntranet then
JStemplateVMENU=replace(JStemplateVMENU,"</ul></div><div id='menuI'><ul id='QS_intranetmenulist'>","<li><a href=""default.asp?iId=" & encrypt(getIntranetHomePage.iId) & """>" & quotrep(customer.intranetName) & "</a><ul>",1,-1,1) & "</li></ul>"
else
JStemplateVMENU=replace(JStemplateVMENU,"</ul></div><div id='menuI'><ul id='QS_intranetmenulist'>","",1,-1,1)' & "</li></ul>"
end if
end if
JStemplateVMENU=replace(JStemplateVMENU,"<div id='menu'><ul id='QS_menulist'>","<ul class=""sf-menu sf-vertical"" id=""SF_vermenu"">",1,-1,1)
JStemplateVMENU=replace(JStemplateVMENU,"<div>","",1,-1,1)
JStemplateVMENU=replace(JStemplateVMENU,"</div>","",1,-1,1)
JStemplateVMENU=JSremoveIDs(JStemplateVMENU)
end function
function JSremoveIDs(value)
dim startpos,endpos
while instr(value,"id='")<>0
startpos=instr(value,"id='")
if startpos<>0 then
endpos=instr(startpos+4,value,"'")
if endpos<>0 then
value=left(value,startpos-1) & right(value,len(value)-endpos)
end if
end if
wend
JSremoveIDs=value
end function
function JStemplateHMENU(p)

	JStemplateHMENU=showMenu
	
	if instr(JStemplateHMENU,"QS_intranetmenulist")<>0 then
	
		if logon.authenticatedIntranet then
			JStemplateHMENU=replace(JStemplateHMENU,"</ul></div><div id='menuI'><ul id='QS_intranetmenulist'>","<li><a href=""default.asp?iId=" & encrypt(getIntranetHomePage.iId) & """>" & quotrep(customer.intranetName) & "</a><ul>",1,-1,1) & "</li></ul>"
		else
			JStemplateHMENU=replace(JStemplateHMENU,"</ul></div><div id='menuI'><ul id='QS_intranetmenulist'>","",1,-1,1)' & "</li></ul>"
		end if
	end if
	
	JStemplateHMENU=replace(JStemplateHMENU,"<div id='menu'><ul id='QS_menulist'>","<ul class=""sf-menu"" id=""SF_hormenu" & p & """>",1,-1,1)
	JStemplateHMENU=replace(JStemplateHMENU,"<div>","",1,-1,1)
	JStemplateHMENU=replace(JStemplateHMENU,"</div>","",1,-1,1)
	JStemplateHMENU=replace(JStemplateHMENU,"id='qs_activelink'","class='activelink'",1,-1,1)	
	
	JStemplateHMENU=JSremoveIDs(JStemplateHMENU)	
	
end function



function qs_artisteer_menu(bInt)
dim menu
set menu=new cls_menu 
menu.cssType="artmenu"
if not bInt then
qs_artisteer_menu=menu.getMenu(null) 
else
qs_artisteer_menu=menu.getIntranetMenu(null) 
end if
set menu=nothing 
qs_artisteer_menu=setActiveLinkWithString(qs_artisteer_menu,"active",null)
qs_artisteer_menu=treatConstants(qs_artisteer_menu,true) 
end function
function qs_artisteer_full_menu(root)
on error resume next
dim menu
set menu=new cls_menu 
if not isNull(root) then
menu.iSubMenuRoot=root.iId
end if
menu.cssType="artmenu"
qs_artisteer_full_menu=menu.getMenu(root)  
if not isNull(root) then
qs_artisteer_full_menu=setActiveLinkWithString(qs_artisteer_full_menu,"active",root.iId)
else
qs_artisteer_full_menu=setActiveLinkWithString(qs_artisteer_full_menu,"active",null)
end if
if customer.intranetUse then
dim gihpi
gihpi=encrypt(getIntranetHomePage.iId)
if logon.authenticatedIntranet then
dim repString
repString="<li><a "
if selectedPage.bIntranet then
repString=repString & " class='active' "
end if
if not isLeeg(getIntranetHomePage.sUserfriendlyURL) and customer.bUserfriendlyURL then
repString=repString& " href="""&getIntranetHomePage.sUserfriendlyURL&""" style=""cursor: default;"" onclick=""javascript:return false;"" id='QS_VMENU" & gihpi & "'>" & getParentLinkForArtisteer(customer.intranetName) & "</a><ul>"
else
repString=repString& " href='#' style=""cursor: default;"" onclick=""javascript:return false;"" id='QS_VMENU" & gihpi & "'>" & getParentLinkForArtisteer(customer.intranetName) & "</a><ul>"
end if
qs_artisteer_full_menu=qs_artisteer_full_menu&menu.getIntranetMenu(null)
qs_artisteer_full_menu=replace(qs_artisteer_full_menu,"</ul><ul class='artmenu'>",repString,1,-1,1) & "</li></ul>"
else
qs_artisteer_full_menu=left(qs_artisteer_full_menu,len(qs_artisteer_full_menu)-5)
qs_artisteer_full_menu=qs_artisteer_full_menu & "<li><a href='default.asp?iId="& gihpi &"' id='QS_VMENU" & gihpi & "'>" & getParentLinkForArtisteer(customer.intranetName) & "</a></li></ul>"
end if
if convertGetal(selectedPage.iId)=0 or request("pageAction")=cloginIntranet then
qs_artisteer_full_menu=setActiveLinkWithString(qs_artisteer_full_menu,"active",null)
end if
end if
set menu=nothing 
qs_artisteer_full_menu=treatConstants(qs_artisteer_full_menu,true) 
qs_artisteer_full_menu=replace(qs_artisteer_full_menu,"class='subcontainer'","",1,-1,1)
err.Clear 
on error goto 0
end function
function qs_mainmenu
dim menu
set menu=new cls_menu
menu.cssType=0
qs_mainmenu=menu.getMenu(null)
set menu=nothing
qs_mainmenu=treatConstants(qs_mainmenu,true)
qs_mainmenu=setActiveLink(qs_mainmenu)
end function
function qs_intranetmenu
dim menu
set menu=new cls_menu
menu.cssType=0
qs_intranetmenu=menu.getIntranetMenu(null)
set menu=nothing
qs_intranetmenu=treatConstants(qs_intranetmenu,true)
qs_intranetmenu=setActiveLink(qs_intranetmenu)
end function
function setActiveLink(sString)
'On Error Resume Next
if convertGetal(selectedPage.iId)<>0 then
if customer.bUserFriendlyURL and not isLeeg(selectedPage.sUserFriendlyURL) then
setActiveLink=replace(sString,"href='" & selectedPage.sUserFriendlyURL & "'","id='qs_activelink' href='" & selectedPage.sUserFriendlyURL & "'",1,-1,1)
else
setActiveLink=replace(sString,"href='default.asp?iId="& encrypt(selectedPage.iId) &"'","id='qs_activelink' href='default.asp?iId="& encrypt(selectedPage.iId) &"'",1,-1,1)
end if
else
if lcase(convertStr(request("pageAction")))=cProfile then
setActiveLink=replace(sString,"href='default.asp?pageAction=profile'","id='qs_activelink' href='default.asp?pageAction=profile'",1,-1,1)
elseif lcase(convertStr(request("pageAction")))="editsite" then
setActiveLink=replace(sString,"href='default.asp?pageAction=editsite'","id='qs_activelink' href='default.asp?pageAction=editsite'",1,-1,1)
else
setActiveLink=sString
end if
end if
setActiveLink=replace(setActiveLink,"class='subcontainer' id='qs_activelink'","class='qs_activelinkandsubcontainer'",1,-1,1)
err.Clear
On error goto 0
end function
function setActiveLinkWithString(sString,sCSSClass,rootID)
On Error Resume Next
dim copySelectedPage, parentID
set copySelectedPage=selectedPage
while convertGetal(selectedPage.iParentID)<>0 and convertGetal(selectedPage.iParentID)<>convertGetal(rootID)
parentID=selectedPage.iParentID
set selectedPage=nothing
set selectedPage=new cls_page
selectedPage.pick(parentID)
wend 
if convertGetal(selectedPage.iId)<>0 then
setActiveLinkWithString=replace(sString,"id='QS_VMENU"&encrypt(selectedPage.iId)&"'","class='" & sCSSClass & "' id='QS_VMENU"& generatePassword & encrypt(selectedPage.iId)&"'",1,-1,1)
else
select case request("pageAction")
case cProfile,cloginIntranet,cForgotPW,cRegister,cWelcome,"editsite"
dim gihpi
gihpi=encrypt(getIntranetHomePage.iId)
setActiveLinkWithString=replace(sString,"id='QS_VMENU" & gihpi & "'","class='" & sCSSClass & "' id='QS_VMENU" & gihpi & "'",1,-1,1)
case else
setActiveLinkWithString=sString
end select
end if
set selectedPage=copySelectedPage
err.Clear
On error goto 0
end function
function art31MENUFIX(sMenu)
art31MENUFIX=replace(sMenu,"<span class=""l""></span>","",1,-1,1)
art31MENUFIX=replace(art31MENUFIX,"<span class=""r""></span>","",1,-1,1)
end function
function art3XVMENUFIX(sMenu)
dim gihpi
if convertGetal(selectedPage.iId)<>0 then
gihpi=encrypt(selectedPage.iId)
sMenu=replace(sMenu,"id='QS_VMENU" & gihpi & "'","class='active' id='QS_VMENU" & gihpi & "'",1,-1,1)
sMenu=replace(sMenu,">" & selectedPage.sTitle & "</span></a><ul>",">" & selectedPage.sTitle & "</span></a><ul class='active'>",1,-1,1)
sMenu=replace(sMenu,">" & selectedPage.sTitle & "</a><ul>",">" & selectedPage.sTitle & "</a><ul class='active'>",1,-1,1)
dim copySelectedPage,parentID,counting
counting=0
set copySelectedPage=selectedPage
while convertGetal(selectedPage.iParentID)<>0 and counting<5 'and convertGetal(selectedPage.iParentID)<>convertGetal(rootID)
sMenu=replace(sMenu,">" & selectedPage.parentPage.sTitle & "</span></a><ul>",">" & selectedPage.parentPage.sTitle & "</span></a><ul class='active'>",1,-1,1)
sMenu=replace(sMenu,">" & selectedPage.parentPage.sTitle & "</a><ul>",">" & selectedPage.parentPage.sTitle & "</a><ul class='active'>",1,-1,1)
parentID=selectedPage.iParentID
set selectedPage=nothing
set selectedPage=new cls_page
selectedPage.pick(parentID)
counting=counting+1
wend 
if convertBool(selectedPage.bIntranet) then
sMenu=replace(sMenu,customer.intranetName & "</span></a><ul>",customer.intranetName & "</span></a><ul class='active'>",1,-1,1)
end if
end if
on error resume next
select case request("pageAction")
case cProfile,cloginIntranet,cForgotPW,cRegister,cWelcome,"editsite"
sMenu=replace(sMenu,customer.intranetName & "</span></a><ul>",customer.intranetName & "</span></a><ul class='active'>",1,-1,1)
sMenu=replace(sMenu,"pageAction=profile'","pageAction="&cProfile&"' class='active'",1,-1,1)
sMenu=replace(sMenu,"pageAction=editsite'","pageAction=editsite' class='active'",1,-1,1)
sMenu=replace(sMenu,customer.intranetName & "</span></a><ul>",customer.intranetName& "</span></a><ul class='active'>",1,-1,1)
gihpi=encrypt(getIntranetHomePage.iId)
sMenu=replace(sMenu,"class='active' id='QS_VMENU" & gihpi & "'","id='QS_VMENU" & gihpi & "'",1,-1,1)
sMenu=replace(sMenu,"onclick=""javascript:return false;"" id='QS_VMENU" & gihpi & "'>","class='active' onclick=""javascript:return false;"" id='QS_VMENU" & gihpi & "'>",1,-1,1)
end select
on error goto 0
art3XVMENUFIX=sMenu
set selectedPage=copySelectedPage
end function%>
