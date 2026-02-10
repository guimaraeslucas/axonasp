<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bForms%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Forms)%><%dim ff
set ff=new cls_FormField
select case request("btnaction")
case "MoveUP"
checkCSRF()
ff.moveUp()
case "MoveDOWN"
checkCSRF()
ff.moveDown()
end select
dim form, fields
set form=new cls_form
set fields=form.fields
dim fixedFieldTypeList
set fixedFieldTypeList=new cls_formFieldTypeList%><p align=center><b><%=form.sName%></b></p><table align=center><tr><td><a href="bs_formFieldEdit.asp?iFormID=<%=encrypt(form.iId)%>"><%=l("newfield")%></a></td><td><%=getIconPP(l("preview"),"search","bs_formPreview.asp?iFormID="&encrypt(form.iId),"","preview"&form.iId,"class=""QSPP""")%></td></tr></table><%if fields.count>0 then%><table align=center id=formFields class=sortable><tr><th><%=l("name")%></th><th><%=l("type")%></th><th><%=l("mandatory")%></th><th><%=l("order")%></th><th>iId</th><th>&nbsp;</th></tr><%dim ffKey
for each ffKey in fields%><tr><td><a href="bs_formFieldEdit.asp?iFormFieldID=<%=encrypt(ffKey)%>"><%=fields(ffKey).sName%></a></td><td><%=fixedFieldTypeList.showSelected("single",fields(ffKey).sType)%></td><td align=center><%=fields(ffKey).isMandatory%></td><td align=center><%=fields(ffKey).iRang%></td><td align=center><%=ffKey%></td><td align=center><table cellpadding=1 cellspacing=0><tr><td><%=getIcon(l("up"),"up","bs_formFields.asp?"&QS_secCodeURL&"&amp;iFormID="&encrypt(form.iId)&"&amp;iFormFieldID="&encrypt(ffKey)&"&amp;btnaction=MoveUP","","up"&ffKey)%></td><td><%=getIcon(l("down"),"down","bs_formFields.asp?"&QS_secCodeURL&"&amp;iFormID="&encrypt(form.iId)&"&amp;iFormFieldID="&encrypt(ffKey)&"&amp;btnaction=MoveDOWN","","down"&ffKey)%></td></tr></table></td></tr><%next%></table><%end if%><!-- #include file="bs_backForm.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
