<!-- #include file="begin.asp"-->


<!-- #include file="beginClient.asp"--><!-- #include file="ad_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><!-- #include file="ad_adminMenu.asp"--><%dim clientproduct,list,i,total,calc,customerList,cList
set clientproduct=new cls_clientproduct
list=clientproduct.list()
total=0
dim cl
set customerList=new cls_client
customerList=customerList.listAll
set cList=server.createObject("scripting.dictionary")
for cl=lbound(customerList,2) to ubound(customerList,2)
cList.Add convertGetal(customerList(0,cl)),customerList(1,cl)
next
if not isNull(list) then%><table class=sortable id=t align=center cellpadding=4 cellspacing=0 style="border-style:none"><tr><td><b>Nr</b></td><td><b>Id</b></td><td><b>Name</b></td><td><b>Renew every</b></td><td><b>Last renewal</b></td><td><b>End service</b></td><td><b>Fee</b></td></tr><%dim cssP
for i=lbound(list,2) to ubound(list,2)
if convertGetal(list(3,i))<=0 then
cssP=""
elseif not isleeg(list(6,i)) then
cssP=""
elseif isLeeg(list(5,i)) then
cssP=";background-color:Yellow"
elseif dateAdd("m",convertGetal(list(2,i)),list(5,i))<dateAdd("d",14,date()) then
cssP=";background-color:Orange"
else
cssP=""
end if%><tr><td style="border-top:1px solid #DDD;text-align:center<%=cssP%>"><%=i+1%></td><td style="border-top:1px solid #DDD;text-align:center<%=cssP%>"><%=list(0,i)%></td><td style="border-top:1px solid #DDD<%=cssP%>"><strong><a href="ad_clientproduct.asp?iClientProductID=<%=encrypt(list(0,i))%>&amp;iClientID=<%=encrypt(list(4,i))%>"><%=list(1,i)%></a></strong><br /><small><%=cList(convertGetal(list(4,i)))%></small></td><td style="border-top:1px solid #DDD;text-align:center<%=cssP%>"><%=list(2,i)%> months</td><td style="border-top:1px solid #DDD;text-align:center<%=cssP%>"><%=convertEuroDate(list(5,i))%></td><td style="border-top:1px solid #DDD;text-align:center<%=cssP%>"><%=convertEuroDate(list(6,i))%></td><td style="border-top:1px solid #DDD;text-align:center<%=cssP%>"><%=list(3,i)%></td></tr><%if convertGetal(list(2,i))=0 or not isleeg(list(6,i)) then
calc=0
else
calc=convertGetal(list(3,i)) * (12/list(2,i))
end if
total=total+calc
next%></table><p align=center>Total: <%=total%></p><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
