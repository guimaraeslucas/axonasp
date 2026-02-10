
<%class cls_fullSearch
public pattern
private re
private rs
sub class_initialize
On Error Resume Next
set re=new regexp
re.Global=true
re.IgnoreCase=true
On Error Goto 0
end sub
sub class_terminate
set re=nothing
end sub
public function customerfields
dim cf, c, counter
set customerfields=server.CreateObject ("scripting.dictionary")
set rs=db.execute("select * from tblCustomer where iId="&cId)
counter=1
dim bGeneral,bBanner,sTopheader,bIntranet,bfooter,adminOSM,adminEM
bGeneral=false
bBanner=false
sTopheader=false
bIntranet=false
bfooter=false
adminOSM=false
adminEM=false
while not rs.eof
for each c in rs.fields
'Response.Write c.name
if testPattern(c) then
select case lcase(c.name)
'general setup
case "sitename","surl","sitetitle","googleanalytics","keywords","sdescription","webmaster","webmasteremail","copyright","sdefaultrsslink"
if not bGeneral then
customerfields.Add counter,"<a target='_blank' href='bs_admin.asp'>"&l("general")&"</a>"
bGeneral=true
end if
case "sbannermenu","bannerapplication","sleftbanner","srightbanner"
if not bBanner then
customerfields.Add counter,"<a target='_blank' href='bs_editBannerMenu.asp'>"&l("bannerunderleftmenu")&"</a>"
bBanner=true
end if
case "stopheader"
if not sTopheader then
customerfields.Add counter,"<a target='_blank' href='bs_logo.asp'>"&l("additionalhtml")&"</a>"
sTopheader=true
end if
case "intranetname","intranetpwemail","intranetmyprofile","slabelregister","semailnewregistrations"
if not bIntranet then
customerfields.Add counter,"<a target='_blank' href='bs_adminIntranet.asp'>"&l("general")&" Intranet</a>"
bIntranet=true
end if
case "sfooter"
if not bfooter then
customerfields.Add counter,"<a target='_blank' href='bs_editFooter.asp'>"&l("footer")&"</a>"
bfooter=true
end if
case "sexplticket","swelcomemessage","sexplprofile"
if not adminOSM then
customerfields.Add counter,"<a target='_blank' href='bs_adminIntranetOSM.asp'>"&l("intranet")&"/"&l("onscreenmessages")&"</a>"
adminOSM=true
end if
case "ssubjectmailticket","smailwelcome","smailticket","intranetpwemail"
if not adminEM then
customerfields.Add counter,"<a target='_blank' href='bs_adminIntranetEM.asp'>"&l("intranet")&"/"&l("automatedemails")&"</a>"
adminEM=true
end if
end select
'customerfields.Add c.name,
counter=counter+1
end if
next 
rs.movenext
wend
set rs=nothing
end function
public function pages
dim page
set pages=server.CreateObject ("scripting.dictionary")
set rs=db.execute("select * from tblPage where (bDeleted="&getSQLBoolean(false)&" or bDeleted is null) and iCustomerID="&cId & " order by sTitle asc")
while not rs.eof
if scanForPattern(rs.fields) then
set page=new cls_page
page.pick(rs("iId"))
pages.Add page.iId,page
set page=nothing
end if
rs.movenext
wend
set rs=nothing
end function
public function constants
dim constant
set constants=server.CreateObject ("scripting.dictionary")
set rs=db.execute("select * from tblConstant where iCustomerID="&cId & " order by sConstant asc")
while not rs.eof
if scanForPattern(rs.fields) then
set constant=new cls_constant
constant.pick(rs("iId"))
constants.Add constant.iId,constant
set constant=nothing
end if
rs.movenext
wend
set rs=nothing
end function
public function forms
dim form
set forms=server.CreateObject ("scripting.dictionary")
set rs=db.execute("select * from tblForm where iCustomerID="&cId & " order by sName")
while not rs.eof
if scanForPattern(rs.fields) then
set form=new cls_form
form.pick(rs("iId"))
forms.Add form.iId,form
set form=nothing
end if
rs.movenext
wend
set rs=nothing
end function
public function formFields
dim formField
set formFields=server.CreateObject ("scripting.dictionary")
set rs=db.execute("SELECT tblFormField.* FROM tblForm INNER JOIN tblFormField ON tblForm.iId = tblFormField.iFormID WHERE tblForm.iCustomerID="& cid & " order by tblFormField.sName")
while not rs.eof
if scanForPattern(rs.fields) then
set formField=new cls_formField
formField.pick(rs("iId"))
formFields.Add formField.iId,formField
set formField=nothing
end if
rs.movenext
wend
set rs=nothing
end function
public function templates
dim template
set templates=server.CreateObject ("scripting.dictionary")
set rs=db.execute("select * from tblTemplate where iCustomerID="&cId & " order by sName")
while not rs.eof
if scanForPattern(rs.fields) then
set template=new cls_template
template.pick(rs("iId"))
templates.Add template.iId,template
set template=nothing
end if
rs.movenext
wend
set rs=nothing
end function
public function feeds
dim feed
set feeds=server.CreateObject ("scripting.dictionary")
set rs=db.execute("select * from tblFeed where iCustomerID="&cId & " order by sName")
while not rs.eof
if scanForPattern(rs.fields) then
set feed=new cls_feed
feed.pick(rs("iId"))
feeds.Add feed.iId,feed
set feed=nothing
end if
rs.movenext
wend
set rs=nothing
end function
public function polls
dim poll
set polls=server.CreateObject ("scripting.dictionary")
set rs=db.execute("select * from tblPoll where iCustomerID="&cId & " order by sCode")
while not rs.eof
if scanForPattern(rs.fields) then
set poll=new cls_poll
poll.pick(rs("iId"))
polls.Add poll.iId,poll
set poll=nothing
end if
rs.movenext
wend
set rs=nothing
end function
public function galleries
dim gallery
set galleries=server.CreateObject ("scripting.dictionary")
set rs=db.execute("select * from tblGallery where iCustomerID="&cId & " order by sName")
while not rs.eof
if scanForPattern(rs.fields) then
set gallery=new cls_gallery
gallery.pick(rs("iId"))
galleries.Add gallery.iId,gallery
set gallery=nothing
end if
rs.movenext
wend
set rs=nothing
end function
public function themes
dim theme
set themes=server.CreateObject ("scripting.dictionary")
set rs=db.execute("select * from tblTheme where iCustomerID="&cId & " order by sName")
while not rs.eof
if scanForPattern(rs.fields) then
set theme=new cls_theme
theme.pick(rs("iId"))
themes.Add theme.iId,theme
set theme=nothing
end if
rs.movenext
wend
set rs=nothing
end function
public function catFields
dim catField
set catFields=server.CreateObject ("scripting.dictionary")
set rs=db.execute("SELECT tblCatalogField.* FROM tblCatalog INNER JOIN tblCatalogField ON tblCatalog.iId = tblCatalogField.iCatalogId WHERE tblCatalog.iCustomerID="& cid &" order by tblCatalogField.sName")
while not rs.eof
if scanForPattern(rs.fields) then
set catField=new cls_catalogField
catField.pick(rs("iId"))
catFields.Add catField.iId,catField
set catField=nothing
end if
rs.movenext
wend
set rs=nothing
end function
public function Imessages
dim Imessage
set Imessages=server.CreateObject ("scripting.dictionary")
set rs=db.execute("SELECT * from tblCustomerIntranetMessage where iCustomerID="&cID)
while not rs.eof
if scanForPattern(rs.fields) then
Imessages.Add "yes","yes"
exit function
end if
rs.movenext
wend
set rs=nothing
end function
private function scanForPattern(row)
On Error Resume Next
scanForPattern=false
dim c
for each c in row
if testPattern(c) then
scanForPattern=true
exit for
end if
next
On Error Goto 0
end function
private function testPattern(value)
On Error Resume Next
testPattern=false
if convertStr(value)<>"" then
re.Pattern=pattern
if re.test(convertStr(value)) then
testPattern=true
end if
end if
On Error Goto 0
end function
public function search
dim op,dKey
dim cpages
set cpages=Pages
dim cconstants
set cconstants=constants
dim cforms
set cforms=forms
dim ctemplates
set ctemplates=templates
dim ccustomerfields
set ccustomerfields=customerfields
dim ccatfields
set ccatfields=catfields
dim cformfields
set cformfields=formfields
dim cfeeds
set cFeeds=feeds
dim cpolls
set cpolls=polls
dim cgalleries
set cgalleries=galleries
dim cImessages
set cImessages=Imessages
dim cThemes
set cThemes=themes
if ccustomerfields.count>0 then
op=op & "<tr><td class=QSlabel valign=top>"&l("setup")&":</td>"
op=op & "<td><ul>"
for each dKey in ccustomerfields
op=op & "<li>" & ccustomerfields(dKey) & "</li>"
next
op=op & "</ul></td></tr>"
end if
if cpages.count>0 then
op=op & "<tr><td class=QSlabel valign=top>"&l("pagelist")&":</td>"
op=op & "<td><ul>"
for each dKey in pages
op=op & "<li>" & cpages(dKey).getClickLink(true) & "</li>"
next
op=op & "</ul></td></tr>"
end if
if cconstants.count>0 then
op=op & "<tr><td class=QSlabel valign=top>"&l("constants")&":</td>"
op=op & "<td><ul>"
for each dKey in cconstants
op=op & "<li><a target='_blank' href='bs_constantEdit.asp?iContentID=" & encrypt(dKey) & "'>" & cconstants(dKey).sConstant & "</a></li>"
next
op=op & "</ul></td></tr>"
end if
if cforms.count>0 then
op=op & "<tr><td class=QSlabel valign=top>"&l("forms")&":</td>"
op=op & "<td><ul>"
for each dKey in cforms
op=op & "<li><a target='_blank' href='bs_formEdit.asp?iFormId=" & encrypt(dKey) & "'>"& cforms(dKey).sname &"</a></li>"
next
op=op & "</ul></td></tr>"
end if
if cformfields.count>0 then
op=op & "<tr><td class=QSlabel valign=top>"&l("form") & " " & l("fields")&":</td>"
op=op & "<td><ul>"
for each dKey in cformfields
op=op & "<li><a target='_blank' href='bs_formFieldEdit.asp?iFormFieldID=" & encrypt(dKey) & "'>"& cformfields(dKey).sName &"</a></li>"
next
op=op & "</ul></td></tr>"
end if
if ctemplates.count>0 then
op=op & "<tr><td class=QSlabel valign=top>"&l("templates")&":</td>"
op=op & "<td><ul>"
for each dKey in ctemplates
op=op & "<li><a target='_blank' href='bs_templateEdit.asp?iTemplateId=" & encrypt(dKey) & "'>"& ctemplates(dKey).sname &"</a></li>"
next
op=op & "</ul></td></tr>"
end if
if cFeeds.count>0 then
op=op & "<tr><td class=QSlabel valign=top>"&l("feed")&":</td>"
op=op & "<td><ul>"
for each dKey in cFeeds
op=op & "<li><a target='_blank' href='bs_feedEdit.asp?iFeedId=" & encrypt(dKey) & "'>"& cFeeds(dKey).sname &"</a></li>"
next
op=op & "</ul></td></tr>"
end if
if cPolls.count>0 then
op=op & "<tr><td class=QSlabel valign=top>Poll:</td>"
op=op & "<td><ul>"
for each dKey in cPolls
op=op & "<li><a target='_blank' href='bs_pollEdit.asp?iPollId=" & encrypt(dKey) & "'>"& cPolls(dKey).sQuestion &"</a></li>"
next
op=op & "</ul></td></tr>"
end if
if cGalleries.count>0 then
op=op & "<tr><td class=QSlabel valign=top>"&l("gallery")&":</td>"
op=op & "<td><ul>"
for each dKey in cGalleries
op=op & "<li><a target='_blank' href='bs_galleryEdit.asp?iGalleryId=" & encrypt(dKey) & "'>"& cGalleries(dKey).sname &"</a></li>"
next
op=op & "</ul></td></tr>"
end if
if cThemes.count>0 then
op=op & "<tr><td class=QSlabel valign=top>"&l("themes")&":</td>"
op=op & "<td><ul>"
for each dKey in cThemes
op=op & "<li><a target='_blank' href='bs_ThemeEdit.asp?iThemeID=" & encrypt(dKey) & "'>"& cThemes(dKey).sname &"</a></li>"
next
op=op & "</ul></td></tr>"
end if
if ccatfields.count>0 then
op=op & "<tr><td class=QSlabel valign=top>"&l("catalog") & " " & l("fields")&":</td>"
op=op & "<td><ul>"
for each dKey in ccatfields
op=op & "<li><a target='_blank' href='bs_catalogFieldEdit.asp?iFieldId=" & encrypt(dKey) & "'>"& ccatfields(dKey).sName &"</a></li>"
next
op=op & "</ul></td></tr>"
end if
if cImessages.count>0 then
op=op & "<tr><td class=QSlabel valign=top>"&l("setup")&":</td>"
op=op & "<td><ul>"
op=op & "<li><a target='_blank' href='bs_adminIntranetEM.asp'>"&l("automatedemails")&"</li>"
op=op & "</ul></td></tr>"
end if
if not isLeeg(op) then
search="<table align=center width=500>"&op&"</table>"
end if
set cpages=nothing
set cconstants=nothing
set cforms=nothing
set ctemplates=nothing
set ccustomerfields=nothing
set ccatfields=nothing
set cformfields=nothing
set cfeeds=nothing
set cpolls=nothing
set cGalleries=nothing
set cImessages=nothing
set cThemes=nothing
end function
end class%>
