
<%pageTitle=customer.sLabelEditSite
dim rsES, sqlES
sqlES="select distinct tblPage.iId, tblPage.sTitle from tblPage "
sqlES=sqlES & " where tblPage.iId in "
sqlES=sqlES & "(select iBodyID from tblContactPage where iContactID="& logon.contact.iId & ")"
sqlES=sqlES & " OR tblPage.iId in "
sqlES=sqlES & "(select iTitleID from tblContactPage where iContactID="& logon.contact.iId & ")"
sqlES=sqlES & " OR tblPage.iId in "
sqlES=sqlES & "(select iLPID from tblContactPage where iContactID="& logon.contact.iId & ")"
sqlES=sqlES & " order by tblPage.sTitle"
set rsES=db.execute(sqlES)
dim getTper,getBper,getLper
set getTper=logon.contact.getTper
set getBper=logon.contact.getBper
set getLper=logon.contact.getLper
dim editMenu, iPageID, bCanEdit, editPageObj
if not rsES.eof then
editMenu="<ul>"
while not rsES.eof
bCanEdit=false
iPageID=rsES(0)
set editPageObj=new cls_page
editPageObj.pick(iPageID)
if getTper.exists(iPageID) then
bCanEdit=true
elseif  getBper.exists(iPageID) then
bCanEdit=true 
elseif  getLper.exists(iPageID) then
bCanEdit=true 
end if
if bCanEdit then 
editMenu=editMenu & "<li><strong><a target='"&encrypt(iPageID)&"' href='default.asp?iId=" & encrypt(iPageID) & "'>" & sanitize(rsES(1)) & "</a></strong>"
if getTper.exists(iPageID) or getBper.exists(iPageID) then
editMenu=editMenu & "&nbsp;|&nbsp;<a class=""bPopupFullWidthReload"" href=""" & C_DIRECTORY_QUICKERSITE & "/asp/"
if isLeeg(editPageObj.sOrderBy) then
editMenu=editMenu & "fs_editPage.asp"
else
editMenu=editMenu & "fs_editListPage.asp"
end if
editMenu=editMenu & "?iId=" & encrypt(iPageID) & """>" & l("modify") & "</a>"
end if
if getLper.exists(iPageID) then
if not isLeeg(editPageObj.sOrderBy) then
editMenu=editMenu &"&nbsp;|&nbsp;<a href=""" & C_DIRECTORY_QUICKERSITE & "/asp/fs_editListItems.asp?iId=" & encrypt(iPageID) & """ class=""bPopupFullWidthReload"">" & l("managearticles") & "</a>&nbsp;"
end if
end if
editMenu=editMenu &"&nbsp;|&nbsp;<a target="""&encrypt(iPageID)&""" href=""default.asp?fpv=1&amp;iId=" & encrypt(editPageObj.iId) & """>Preview</a>"
editMenu=editMenu & "</li>"
end if
set editPageObj=nothing
rsES.movenext
wend 
editMenu=editMenu& "</ul>"
end if
if bCanEdit then
Session(cId & "isAUTHENTICATEDasUSER")=true
end if
pageBody=editMenu%>
