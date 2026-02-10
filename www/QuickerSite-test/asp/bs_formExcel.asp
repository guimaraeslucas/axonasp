<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bFormExport%><%dim form, submissions, sKey, submission, values, fields, cField, showCatalog
showCatalog=false
set form=new cls_form
set fields=form.fields
set submissions=form.submissions
for each sKey in  submissions
if not isLeeg(submissions(sKey).iItemID) then
showCatalog=true
exit for
end if
next
dim table
table="<table><tr><td style='background-color:Gainsboro;'><b>ID</b></td><td style='background-color:Gainsboro;'><b>" & l("date") &"&nbsp;&amp;&nbsp;"& l("hour") & "</b></td>"
if showCatalog then
table=table& "<td style='background-color:Gainsboro;'><b>" & l("catalog") & "</b></td>"
table=table& "<td style='background-color:Gainsboro;'><b>" & l("item") & "</b></td>"
end if
for each cField in fields
if fields(cField).sType<>sb_ff_comment then
table=table& "<td style='background-color:Gainsboro;'><b>" & cleanUpStr(fields(cField).sName) & "</b></td>"
end if
next
table=table& "</tr>"
dim tableRow
for each sKey in  submissions
tableRow=""
set submission=submissions(sKey)
set values=submission.values(fields)
tableRow=tableRow& "<tr><td>" & sKey & "</td>"
tableRow=tableRow& "<td>" & formatTimeStamp(submission.dCreatedTS) &"</td>"
if showCatalog then
tableRow=tableRow& "<td>"& submission.item.catalog.sName & "</td>"
tableRow=tableRow& "<td>"& submission.item.sTitle & "</td>"
end if
for each cField in fields
if fields(cField).sType<>sb_ff_comment then
tableRow=tableRow& "<td>"
select case fields(cField).sType
case sb_ff_checkbox
tableRow=tableRow & convertCheckedYesNo(sanitize(values(cField)))
case sb_ff_file,sb_ff_image
if not isLeeg(values(cField)) then
if isLeeg(fields(cField).sFileLocation) then
tableRow=tableRow & linkUrls(customer.sVDUrl & application("QS_CMS_userfiles") & values(cField))
end if
end if
case else
tableRow=tableRow& sanitize(values(cField))
end select
tableRow=tableRow& "</td>"
end if
next
tableRow=tableRow& "</tr>"
table=table&tableRow
set values=nothing
set submission=nothing
next
table=table & "</table>"
dim excelfile
set excelfile=new cls_excelfile
excelfile.export(table)
excelfile.redirectLink()
set excelfile=nothing %>
