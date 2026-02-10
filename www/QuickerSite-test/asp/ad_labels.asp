<!-- #include file="begin.asp"-->


<!-- #include file="ad_security.asp"--><!-- #include file="includes/commonheader.asp"--><body style="background-color:#FFF"><!-- #include file="bs_initBack.asp"--><%dim languageList, languageTable
set languageList	= new cls_languageListNew
set languageTable	= languageList.table
dim labelList, labelTable
set labelList	= new cls_labelList
set labelTable	= labelList.table%><table align=center><tr><td align=center><a href="ad_label.asp"><%=l("newlabel")%></a></td></tr></table><br /><%if labelTable.count>0 then%><p align=center><%=labelTable.count%>&nbsp;<%=l("labels")%></p><table align=center cellpadding=3 border=1 cellspacing=0 class=sortable id=labellist><tr><th>ID</th><th>Code</th><%dim langKey, labelKey
for each langKey in languageTable%><th><%=languageTable(langKey).sLanguage%></th><%next%><th>&nbsp;</th></tr><%for each labelKey in labelTable%><tr><td valign=top><%=labelKey%></td><td valign=top style="height:30px"><a name=l<%=labelKey%>></a><%=labelTable(labelKey).sCode%></td><%for each langKey in languageTable%><td valign=top><%=labelTable(labelKey).values(langKey)%></td><%next%><td valign=top><a href="ad_label.asp?iId=<%=labelKey%>"><%=l("modify")%></a></td></tr><%next%></table><br /><%end if
set labelTable=nothing
set labelList=nothing
set languageList=nothing
set languageTable=nothing%><!-- #include file="ad_back.asp"--><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><!-- #include file="bs_endBack.asp"--><%=message.showAlert()%><%=cPopup.dumpJS()%> 
</body></html><%cleanUPASP%>
