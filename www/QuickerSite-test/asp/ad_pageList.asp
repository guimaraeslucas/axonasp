<!-- #include file="begin.asp"-->


<!-- #include file="ad_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%dim search
set search=new cls_search
search.showAll=convertBool(Request.QueryString ("showAll"))
dim allPages, i
allPages=search.allPages%><br /><%if search.showAll then%><p align=center><%=l("totalnumberofpages")%>: <b><%=convertGetal(search.recordCount)%></b></p><%end if%><%if not isNull(allPages) and search.recordCount>0 then%><table id=allPages class=sortable align=center cellspacing=0 cellpadding=3><tr><th><%=l("title")%></th><th><%=l("account")%></th><th><%=l("createdon")%></th><th><%=l("updatedon")%></th><th><%=l("visits/day")%></th></tr><%for i=lbound(allPages,2) to ubound(allPages,2)%><tr><td style="border-top:1px solid #DDD"><a target="n<%=allPages(0,i)%>" href="<%=replace(allPages(5,i),C_DIRECTORY_QUICKERSITE,"")%><%=C_DIRECTORY_QUICKERSITE%>/default.asp?iID=<%=encrypt(allPages(0,i))%>"><%Response.Write replace(quotrep(left(allPages(1,i),30)),"&","&amp;",1,-1,1)
if len(allPages(1,i))>30 then Response.Write "..."%></a></td><td style="border-top:1px solid #DDD"><%=quotrep(allPages(2,i))%></td><td style="border-top:1px solid #DDD"><%=formatTimeStamp(allPages(3,i))%></td><td style="border-top:1px solid #DDD"><%=formatTimeStamp(allPages(4,i))%></td><td style="border-top:1px solid #DDD" align=center><%=round(convertGetal(allPages(6,i)/aantalDagenWithDate(allPages(7,i))))%></td></tr><%next%></table><%end if%><p align=center><a href="ad_pageList.asp?showall=<%=true%>"><%=l("showall")%></a></p><br /><!-- #include file="ad_back.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
