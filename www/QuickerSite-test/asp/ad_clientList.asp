<!-- #include file="begin.asp"-->


<!-- #include file="beginClient.asp"--><!-- #include file="ad_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><!-- #include file="ad_adminMenu.asp"--><%dim client,list,i
set client=new cls_client
list=client.listAll()
if not isNull(list) then%><table class=sortable id=t align=center cellpadding=4 cellspacing=0 style="border-style:none"><tr><td><b>Nr</b></td><td><b>iId</b></td><td><b>Name</b></td><td><b>Email</b></td></tr><%for i=lbound(list,2) to ubound(list,2)%><tr><td style="border-top:1px solid #DDD;text-align:center"><%=i+1%></td><td style="border-top:1px solid #DDD;text-align:center"><%=list(0,i)%></td><td style="border-top:1px solid #DDD"><strong><a href="ad_client.asp?iClientID=<%=encrypt(list(0,i))%>"><%=quotrep(list(1,i))%></a></strong></td><td style="border-top:1px solid #DDD"><a href="mailto:<%=quotrep(list(3,i))%>"><%=quotrep(list(3,i))%></a></td></tr><%next%></table><%end if
set client=nothing%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
