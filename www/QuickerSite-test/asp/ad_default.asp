<!-- #include file="begin.asp"-->


<!-- #include file="ad_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%dim iDeleteID,delCust
for each iDeleteID in request.form("iDeleteID")
set delCust=new cls_customer
delCust.pick(iDeleteID)
delCust.remove()
set delCust=nothing
next
dim customerList
set customerList=new cls_customerList
if request.querystring("showDetail")="" then
dim getAll,cccc,totsize
cccc=0
totsize=0
set getALL=customerList.getAll%><p align="center"><a href="ad_default.asp?showDetail=true">Show Statistics</a></p><p align="center"><span id="tot">Please wait...</span></p><p align="center"><span id="totsize"></span></p><form action=ad_default.asp method=post><table align=center class=sortable id=main cellpadding=4 cellspacing=0 border=0><tr><th>Nr</th><th>iId</th><th>Date</th><th><%=l("name")%></th><th>&nbsp;</th><th>&nbsp;</th><th>Delete</th></tr><%while not getAll.eof 
cccc=cccc+1
totsize=convertGetal(getAll("iFolderSize"))+totsize
response.write "<tr>"
response.write "<td style=""border-top:1px solid #DDD"">"& cccc &"</td>"
response.write "<td style=""border-top:1px solid #DDD"">"& getAll("iId") &"</td>"
'response.write "<td style=""border-top:1px solid #DDD;text-align:center""><a onclick=""javascript: return confirm('Are you sure to reset this folder size?')"" href='ad_default.asp?iResetID=" & getAll("iId") & "'>"& convertGetal(getAll("iFolderSize")) &"</a></td>"
response.write "<td style=""border-top:1px solid #DDD"">"& convertEurodate(getAll("dCreatedTS")) &"</td>"
response.write "<td style=""border-top:1px solid #DDD""><a href='ad_customer.asp?iId="&getAll("iId")&"'>"& replace(quotrep(getAll("sName")),"&","&amp;",1,-1,1) &"</a><br /><small>" & getAll("sUrl") & "</small></td>"
response.write "<td style=""border-top:1px solid #DDD""><a target='_new' href='" & getAll("sUrl") & " '>Visit</a> | <a href=""mailto:" & getAll("webmasterEmail") & """>Mail</a></td>"
response.write "<td style=""border-top:1px solid #DDD""><a href='" & getBacksiteURL(getAll("sUrl")) &"' target='_new'>Backsite</a></td>"
if getAll("iId")<>cId then
response.write "<td style=""border-top:1px solid #DDD"" align=""center""><input type=""checkbox"" name=""iDeleteID"" value=""" & getAll("iId") & """ /></td>"
else
response.write "<td style=""border-top:1px solid #DDD"" align=""center"">&nbsp;</td>"
end if
response.write "</tr>"
getAll.movenext
wend%></table><p align=center style="width:100%"><input type="submit" onclick="javascript:return confirm('Are you sure to remove the selected websites?');" value="Delete selected" /></p></form><script type="text/javascript">document.getElementById("tot").innerHTML='Total nmbr of sites: <b><%=cccc%></b> sites'</script><script type="text/javascript">document.getElementById("totsize").innerHTML='Total disk space: <b><%=totsize%></b> MB'</script><%set getAll=nothing
else
customerList.showRecentOnly=false
dim customers
set customers=customerList.table
if customers.count>0 then%><table align=center class=sortable id=main cellpadding=5 cellspacing=0 border=0><tr><th>iId</th><th><%=l("name")%></th><th><%=l("allowapplications")%>?</th><th>&nbsp;</th><th>&nbsp;</th><th><%=l("hits/day")%></th><th><%=l("visits/day")%></th></tr><%dim totalHitsPerDag
dim totalBezoekersPerDag
totalHitsPerDag=0
totalBezoekersPerDag=0
dim aantalDagen
dim c, cObj
for each c in customers
set cObj=customers(c)
aantalDagen=cObj.aantalDagen
totalHitsPerDag=totalHitsPerDag+convertGetal(cObj.iTotalHits/aantalDagen)
totalBezoekersPerDag=totalBezoekersPerDag+convertGetal(cObj.iMaxVisits/aantalDagen)%><tr><td style="border-top:1px solid #DDD" align=center width="50"><%=c%></td><td style="border-top:1px solid #DDD"><a href="ad_customer.asp?iId=<%=c%>"><%=cObj.sName%></a></td><td style="border-top:1px solid #DDD" align=center><%=jaNeen(cObj.bApplication)%></td><td style="border-top:1px solid #DDD" align=center width="60"><a target="visit<%=c%>" href="<%=cObj.sUrl%>"><%=l("visit")%></a></td><td style="border-top:1px solid #DDD" align=center width="60"><a target="visit<%=c%>" href="<%=cObj.sQSUrl%>/backsite">Backsite</a></td><td style="border-top:1px solid #DDD" width="60" align=center><%=round(convertGetal(cObj.iTotalHits/aantalDagen),0)%></td><td style="border-top:1px solid #DDD" width="60" align=center><%=round(convertGetal(cObj.iMaxVisits/aantalDagen),0)%></td></tr><%set cObj=nothing
next%></table><p align=center><%=l("totalsforQuickerSiteserver")%>:</p><table align=center><tr><td align=right><b><%=round(totalHitsPerDag,0)%></b></td><td><%=l("hits/day")%></td></tr><tr><td align=right><b><%=round(totalBezoekersPerDag,0)%></b></td><td><%=l("visits/day")%></td></tr></table><%else
Response.Redirect ("ad_customer.asp")
end if
end if%><%if convertBool(Request.QueryString ("newAccount")) then%><script type="text/javascript">alert('<%=l("createdaccount")%>\n\niId: <%=request.querystring("iId")%>');
</script><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
