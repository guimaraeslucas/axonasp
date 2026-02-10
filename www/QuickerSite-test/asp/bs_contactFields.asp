<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bIntranetContacts%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Intranet)%><%=getBOHeaderIntranet(btn_contacts)%><%dim cf
set cf=new cls_contactField
select case request("btnaction")
case "MoveUP"
checkCSRF()
cf.moveUp()
case "MoveDOWN"
checkCSRF()
cf.moveDown()
end select
dim fixedFieldTypeList
set fixedFieldTypeList=new cls_contactFieldTypeList
dim contactFields,contactField
set contactFields=customer.contactFields(false)%><p align=center><a href="bs_contactFieldEdit.asp"><b><%=l("newfield")%></b></a></p><%if contactFields.count>0 then%><table align=center class=sortable id=contactField cellpadding="2"><tr><th><%=l("name")%></th><th><%=l("type")%></th><th><%=l("mandatory")%></th><th><%=l("searchfield")%></th><th><%=l("publicprofile")%></th><th>iId</th><th><%=l("sortorder")%></th><th>&nbsp;</th></tr><%for each contactField in contactFields%><tr><td><a href="bs_contactFieldEdit.asp?iFieldID=<%=encrypt(contactField)%>"><%=contactFields(contactField).sFieldname%></a></td><td align=center><%=fixedFieldTypeList.showSelected("single",contactFields(contactField).sType)%></td><td align=center><%=contactFields(contactField).isMandatory%></td><td align=center><%=contactFields(contactField).isSearchField%></td><td align=center><%=contactFields(contactField).isProfileField%></td><td align=center><%=contactField%></td><td align=center><%=contactFields(contactField).iRang%></td><td><table cellpadding=1 cellspacing=0><tr><td><%=getIcon(l("up"),"up","bs_contactFields.asp?"&QS_secCodeURL&"&amp;iFieldID="&encrypt(contactField)&"&amp;btnaction=MoveUP","","up"&contactField)%></td><td><%=getIcon(l("down"),"down","bs_contactFields.asp?"&QS_secCodeURL&"&amp;iFieldID="&encrypt(contactField)&"&amp;btnaction=MoveDOWN","","down"&contactField)%></td></tr></table></td></tr><%next%></table><br /><%end if%><!-- #include file="bs_backContact.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
