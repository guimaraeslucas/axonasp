<!-- #include file="begin.asp"-->


<!-- #include file="ad_security.asp"--><!-- #include file="includes/commonheader.asp"--><body style="background-color:#FFF"><!-- #include file="bs_initBack.asp"--><%dim languageList, languageTable, langKey
set languageList	= new cls_languageListNew
set languageTable	= languageList.table
dim labellist
set labellist=new cls_labellist
dim label, values
set label=new cls_label
label.pick(request("iId"))
set values=label.values
if Request ("btnaction")=l("save") then
checkCSRF()
label.getRequestValues()
if label.save() then
for each langKey in languageTable
labellist.refresh(langKey)
next
Response.Redirect ("ad_label.asp?fbMessage=fb_saveOK&iId="& label.iId)
end if
end if
if Request ("btnaction")=l("delete") then
checkCSRF()
label.remove()
Response.Redirect ("ad_labels.asp")
end if%><script type="text/javascript">function copyToAll(def) {<%for each langKey in languageTable%>document.getElementById('lk_<%=langKey%>').value=def.value;<%next%>}</script><form action="ad_label.asp" name="mainform" method="post"><%=QS_secCodeHidden%><input type="hidden" name="btnaction" value="<%=l("save")%>" /><input type="hidden" name="iId" value="<%=label.iID%>" /><%if isLeeg(label.iId) then%><p align=center><%=l("label")%>: <input type=text value="<%=quotRep(label.sCode)%>" name="sCode" /></p><%else%><table align=center cellpadding="2"><tr><td align=center><a href="ad_label.asp"><%=l("newlabel")%></a></td></tr></table><p align=center><input type=hidden value="<%=label.sCode%>" name="sCode" /><%=l("label")%>: <%=label.sCode%></p><%end if%><table align=center cellpadding="2"><%dim lrunner
lrunner=0
for each langKey in languageTable
lrunner=lrunner+1%><tr><td width="<%=90/round(languageTable.count,0)%>%"><%=languageTable(langKey).sLanguage%></td><td width="<%=90/round(languageTable.count,0)%>%" valign=top><textarea cols=90 rows=2 id="lk_<%=langKey%>" name="lk_<%=langKey%>"><%=quotRep(values(langKey))%></textarea><%if lrunner=1 then%><a href='#' onclick="javascript:copyToAll(document.getElementById('lk_<%=langKey%>'))">Copy to all</a><%end if%></td></tr><%next%></tr><tr><td colspan="<%=languageTable.count%>" align=center><input class="art-button" type=submit name=dummy value="<%=l("save")%>" />&nbsp;
<%if isNumeriek(label.iId) then%><a href="#" onclick="javascript:if(confirm('<%=l("areyousure")%>')){location.assign('ad_label.asp?<%=QS_secCodeURL%>&amp;btnaction=<%=l("delete")%>&amp;iId=<%=label.iId%>')}"><img alt="<%=l("delete")%>" src="<%=C_DIRECTORY_QUICKERSITE%>/fixedImages/dustbin.gif" border=0 /></a><%end if%></td></tr></table></form><table align=center><tr><td align=center>-> <b><a href="ad_labels.asp"><%=l("back")%></a></b> <-</td></tr></table><%set languageList=nothing
set languageTable=nothing
set values=nothing
set label=nothing%><!-- #include file="bs_endBack.asp"--><%=message.showAlert()%><%=cPopup.dumpJS()%> 
</body></html><%cleanUPASP%>
