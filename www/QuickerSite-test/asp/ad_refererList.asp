<!-- #include file="begin.asp"-->


<!-- #include file="ad_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%dim rs, sql,selectedname
selectedname=Request.QueryString ("selectedname")
sql="SELECT top 15000 tblSession.Referer, tblCustomer.sName, tblCustomer.sUrl, tblSession.dTS, tblCustomer.iId "
sql=sql&" FROM tblCustomer INNER JOIN tblSession ON tblCustomer.iId = tblSession.iCustomerID "
if selectedname<>"" then
sql=sql&" WHERE tblCustomer.sName='" & replace(left(selectedname,100),"'","''",1,-1,1) & "'"
end if
sql=sql&" ORDER BY tblSession.dTS desc"
set rs=db.execute(sql)
if selectedname="" then
dim statsDict,cname
set statsDict=server.CreateObject ("scripting.dictionary")
while not rs.eof
cname=rs(1)
if not statsDict.Exists (cname) then
statsDict.Add cname,1
else
statsDict(cname)=convertGetal(statsDict(cname))+1
end if 
rs.movenext
wend 
SortDictionary2 statsDict,1%><table id=allPages1 class=sortable align=center cellpadding=4 cellspacing=0><tr><th>Account</th><th>Referers</th></tr><%dim cItem
for each cItem in statsDict
Response.Write "<tr><td style=""border-top:1px solid #DDD""><a href='ad_refererlist.asp?selectedname="& server.URLEncode (cItem) &"'>"& cItem & "</td><td  style=""border-top:1px solid #DDD"" align=""center"">" & statsDict(cItem) & "</td></tr>"
next %></table><%else
dim selCounter
selCounter=0%><p align=center>Referers for <b><%=selectedname%></b>:</p><table id=allPages1 class=sortable align=center cellpadding=4 cellspacing=0><tr><th>Referrer</th><th>Date</th></tr><%while not rs.eof 
Response.Write  "<tr><td style=""border-top:1px solid #DDD""><a title=""" & sanitize(rs(0)) & """ target=""_blank"" href=""" & sanitize(rs(0)) & """>" & sanitize(left(rs(0),60)) & "</a></td>"
Response.Write  "<td style=""border-top:1px solid #DDD"">"& convertCalcDate(convertEuroDate(rs(3))) &"</td></tr>"
rs.movenext
selCounter=selCounter+1
wend%></table><p align="center"><%=selCounter%> results</p><%end if
set rs=nothing%><!-- #include file="ad_back.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
