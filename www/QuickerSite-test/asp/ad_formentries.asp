<!-- #include file="begin.asp"-->


<!-- #include file="ad_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%if request.querystring("iSubmissionID")<>"" then
response.clear
dim subm,form,values,v,fields,fa
set subm=new cls_submission
subm.pick(request.querystring("iSubmissionID"))
set form=subm.form
set fields=form.fields
set values=subm.values(fields)
response.write "<ul>"
for each fa in values
response.write "<li><b>" & fields(fa).sName & "</b> - " & quotrep(values(fa)) & "</li>"
next
response.write "</ul>"
dim rsC
set rsC=db.execute("select count(*) from tblFormSubmission where iFormID=" & convertGetal(form.iId))
response.write "<p>There are <b>" & rsC(0) & "</b> similar submissions</p>"
set rsC=nothing
set rsC=db.execute("select count(*) from tblFormSubmission where iFormID in (select iId from tblForm where iCustomerID=" & form.iCustomerID & ")")
response.write "<p>There are <b>" & rsC(0) & "</b> submissions for this total website</p>"
set rsC=nothing
response.end
end if%><table style="border-style:none" align="center" cellpadding="4" cellspacing="0"><tr><td>Date</td><td>Message</td><td>Site</td></tr><%dim rs, sql, rs2, counter
counter=0
sql="SELECT tblFormsubmission.dCreatedTS,tblForm.sName, tblCustomer.sName, tblFormsubmission.iId FROM tblCustomer INNER JOIN (tblForm INNER JOIN tblFormSubmission ON tblForm.iId = tblFormSubmission.iFormID) ON tblCustomer.iId = tblForm.iCustomerID order by tblFormsubmission.dCreatedTS desc"
set rs=db.execute(sql)
while not rs.eof and counter<250
counter=counter+1
response.write "<tr>"
response.write "<td>" & counter & "</td>"
response.write "<td style=""border-bottom:1px solid #DDD"">" & convertEurodate(rs(0)) & "</td>"
response.write "<td style=""border-bottom:1px solid #DDD""><a target=""_new"" class=""QSPP"" href=""ad_formentries.asp?iSubmissionID=" & rs(3) & """>" & rs(1) & "</a></td>"
response.write "<td style=""border-bottom:1px solid #DDD"">" & rs(2) & "</td>"
response.write "</tr>"
rs.movenext
wend
set rs=nothing%></table><!-- #include file="ad_back.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
