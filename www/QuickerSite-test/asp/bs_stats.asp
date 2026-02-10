<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bStats%><!-- #include file="bs_process.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Setup)%><%=getBOSetupMenu(btn_Stats)%><%dim stats, item
set stats=new cls_search
stats.value=""
stats.orderby="iHits desc"
stats.includeURL=false
stats.allowEmptyString=true
stats.includePasswordProtected=true
stats.includeIntranet=true
set stats=stats.results
dim aantalDagen
aantalDagen=customer.aantalDagen%><p align=center><%=l("hitsexpl")%><br /><%=l("visitsexpl")%></p><table align=center cellpadding=4 cellspacing=0 id=stats class=sortable><tr><th><%=l("title")%></th><th><%=l("hits")%>&nbsp;<font color="#ffa500">(RSS)</font></th><th><%=l("hits/day")%>&nbsp;<font color="#ffa500">(RSS)</font></th><th><%=l("visits")%></th><th><%=l("visits/day")%></th></tr><%dim iHitsRSS
for each item in stats
 
iHitsRSS=round(convertGetal(stats(item).iHitsRSS))%><tr><td style="border-top:1px solid #DDD"><%=stats(item).sTitle%></td><td style="border-top:1px solid #DDD" align=center><%=stats(item).iHits%><%if iHitsRSS<>0 then Response.Write "&nbsp;<font color='#ffa500'>(" & round(convertGetal(stats(item).iHitsRSS)) &")</font>"%></td><td style="border-top:1px solid #DDD" align=center><%=round(convertGetal(stats(item).iHits/aantalDagen))%><%if iHitsRSS<>0 then Response.Write "&nbsp;<font color='#ffa500'>(" & round(convertGetal(stats(item).iHitsRSS/aantalDagen)) &")</font>"%></td><td style="border-top:1px solid #DDD" align=center><%=stats(item).iVisitors%></td><td style="border-top:1px solid #DDD" align=center><%=round(convertGetal(stats(item).iVisitors/aantalDagen))%></td></tr><%next%></table><p align=center><%=l("resetted")%>&nbsp;<%=l("on")%>&nbsp;<%=formatTimeStamp(customer.dResetStats)%><br /><a onclick="javascript:return confirm('<%=l("areyousure")%>');" href="?<%=QS_secCodeURL%>&amp;btnaction=ResetStats"><b><%=l("reset")%></b></a><%if customer.bScanreferer then%>&nbsp;-&nbsp;<a href="bs_referers.asp"><b><%=l("referringsites")%></b></a><%end if%></p><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
