<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bFormExport%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Forms)%><%dim form,fields,cField,submission,postback,rs
set form=new cls_form
set fields=form.fields
set submission=new cls_submission
if submission.iFormID<>form.iId then response.end
postback=convertBool(request.form("postback"))
if postback then
on error resume next
for each cField in fields
select case fields(cField).sType
case sb_ff_file,sb_ff_image
'do nothing
case else
set rs=db.getDynamicRS
rs.open "select * from tblFormFieldValue where iSubmissionID="& submission.iId & " and iFormFieldId="& cField
if rs.eof then
rs.addNew()
rs("iFormFieldId")=cField
rs("iSubmissionID")=submission.iId
end if
rs("sValue")=request.form(encrypt(cField))
rs.update
rs.close
set rs=nothing
end select
next
on error goto 0
response.redirect("bs_formExport.asp?iFormID=" & encrypt(form.iId) & "#" & encrypt(submission.iId))
end if%><table align=center><tr><td align=center>-> <b><a href="bs_formExport.asp?iFormID=<%=encrypt(form.iId)%>">Back</a></b> <-</td></tr></table><br /><form action="bs_formSubmissionEdit.asp" method="post" name="editform"><input type="hidden" name="iSubmissionID" value="<%=encrypt(submission.iId)%>" /><input type="hidden" name="iFormID" value="<%=encrypt(form.iId)%>" /><input type="hidden" name="postback" value="<%=true%>" /><table align="center" cellpadding="2"><%on error resume next
for each cField in fields
select case fields(cField).sType
case sb_ff_file,sb_ff_image
'do nothing
case else
response.write "<tr><td class=""label"">" & cleanUpStr(fields(cField).sName) & "</td><td><textarea name=""" & encrypt(cField) & """ cols=""80"" rows=""3"">" & quotrep(submission.values(fields)(cfield)) & "</textarea></td></tr>"
end select
next
on error goto 0
set fields=nothing
set submission=nothing
set form=nothing%><tr><td>&nbsp;</td><td><input class="art-button" type="submit" onclick="javascript:return confirm('<%=quotrepJS(l("areyousure"))%>');" value="<%=l("save")%>" /></td></tr></table></form><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
