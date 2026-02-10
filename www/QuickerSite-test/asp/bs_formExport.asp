<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bFormExport%><!-- #include file="includes/commonheader.asp"--><body style="background-color:#FFF"><%dim form, submissions, sKey, submission, values, fields, cField, showCatalog
showCatalog=false
set form=new cls_form
set fields=form.fields
set submission=new cls_submission
if convertGetal(submission.iId)<>0 then
checkCSRF()
submission.remove(fields)
end if
if Request.QueryString ("pageAction")="removeAll" then
checkCSRF()
form.removeAllSubmissions()
end if
set submissions=form.submissions
if submissions.count>0 then
for each sKey in  submissions
if not isLeeg(submissions(sKey).iItemID) then
showCatalog=true
exit for
end if
next%><table align=center><tr><td><b><%= submissions.count%></b>&nbsp;<%=l("submissions")%>&nbsp;</td><td><%=getIcon("Excel","excel","#","javascript:window.open('bs_formExcel.asp?iFormID="&encrypt(form.iId)&"');","excel")%></td></tr></table><br /><table align=center class=sortable id=jol cellspacing=3><tr><th>&nbsp;</th><th><%=l("date")%>&nbsp;&amp;&nbsp;<%=l("hour")%></th><%if showCatalog then%><th><%=l("catalog")%></th><th><%=l("item")%></th><%end if%><%for each cField in fields
if fields(cField).sType<>sb_ff_comment then%><th><%=cleanUpStr(fields(cField).sName)%></th><%end if
next%><th>&nbsp;</th></tr><%for each sKey in  submissions
set submission=submissions(sKey)
set values=submission.values(fields)%><tr><td valign=top><a name="<%=encrypt(sKey)%>"></a><a href="bs_formSubmissionEdit.asp?iFormID=<%=encrypt(form.iId)%>&amp;iSubmissionID=<%=encrypt(sKey)%>"><%=l("modify")%></a></td><td valign=top><%=formatTimeStamp(submission.dCreatedTS)%></td><%if showCatalog then%><td valign=top><%=submission.item.catalog.sName%></td><td valign=top><%=submission.item.sTitle%></td><%end if%><%for each cField in fields
if fields(cField).sType<>sb_ff_comment then%><td valign=top><%select case fields(cField).sType
case sb_ff_checkbox
Response.Write convertCheckedYesNo(sanitize(values(cField)))
case sb_ff_file,sb_ff_image
if not isLeeg(values(cField)) then
if isLeeg(fields(cField).sFileLocation) then
Response.Write "<a target='" & encrypt(cField) & "' href="& """" & C_VIRT_DIR & application("QS_CMS_userfiles") & values(cField) & """" &">"& cleanUpStr(fields(cField).sName) &"</a>"
end if
end if
case sb_ff_radio
'if fields(cField).bAllowMS then
'	response.write replace (sanitize(values(cField)),",","<br />",1,-1,1)
'else
Response.Write sanitize(values(cField))
'end if
case else
Response.Write sanitize(values(cField))
end select%>&nbsp;</td><%end if
next%><td valign=top><a href="#" onclick="javascript:if(confirm('<%=l("areyousure")%>')){location.assign('bs_formExport.asp?<%=QS_secCodeURL%>&amp;iFormId=<%=encrypt(form.iId)%>&amp;iSubmissionID=<%=encrypt(sKey)%>')};"><img alt="<%=l("delete")%>" src="<%=C_DIRECTORY_QUICKERSITE%>/fixedImages/dustbin.gif" border=0 /></a></td></tr><%set values=nothing
set submission=nothing
next%></table><p align=center><a href="#" onclick="javascript:if(confirm('<%=l("areyousure")%>')){location.assign('bs_formExport.asp?<%=QS_secCodeURL%>&amp;iFormId=<%=encrypt(form.iId)%>&amp;pageAction=removeAll')};"><%=l("removeall")%></a></p><%else%><p align=center><%=l("nodata")%></p><%end if%></body></html><%cleanUPASP%>
