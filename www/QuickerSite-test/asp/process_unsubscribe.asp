
<%if len(request.querystring("e"))=24 then
'get the key
dim rsE
set rsE=db.execute("select * from tblNewsletterCategorySubscriber where iCustomerID=" & cId & " and sKey='" & left(cleanup(request.querystring("e")),24) & "'")
if not rsE.eof then
dim QS_nl
set QS_nl=new cls_newslettercategory
QS_nl.pick(rsE("iCategoryID"))
if convertGetal(QS_nl.iId)<>0 then
dim USbody,UStitle
UStitle=convertStr(QS_nl.sUnsubscribeFBTitle)
UStitle=replace(UStitle,"[NL_NAME]",convertStr(quotrep(rsE("sName"))),1,-1,1)
UStitle=replace(UStitle,"[NL_EMAIL]",convertStr(quotrep(rsE("sEmail"))),1,-1,1)
USbody=convertStr(QS_nl.sUnsubscribeFB)
USbody=replace(USbody,"[NL_NAME]",convertStr(quotrep(rsE("sName"))),1,-1,1)
USbody=replace(USbody,"[NL_EMAIL]",convertStr(quotrep(rsE("sEmail"))),1,-1,1)
pageBody=USbody
pageTitle=UStitle
'remove
if not convertBool(rsE("bActive")) then
set rsE=nothing
redirectToHP()
else
db.execute("update tblNewsletterCategorySubscriber set bActive=" & getSQLBoolean(false) & " where iCustomerID=" & cId & " and sKey='" & left(cleanup(request.querystring("e")),24) & "'")
end if
if not isLeeg(QS_nl.sNotifEmail) then
dim ncEMail
set ncEMail=new cls_mail_message
ncEMail.receiver=QS_nl.sNotifEmail
ncEMail.subject="Unsubscribe from email list '" & QS_nl.sName & "': " & convertStr(quotrep(rsE("sEmail")))
ncEMail.body="email: " & quotrep(rsE("sEmail")) & "<br />" & "name: " & convertStr(quotrep(rsE("sName")))
ncEMail.send
set ncEMail=nothing
end if
set rsE=nothing
else
set rsE=nothing
redirectToHP()
end if
set QS_nl=nothing
else
set rsE=nothing
redirectToHP()
end if
else
redirectToHP()
end if%>
