<!-- #include file="begin.asp"-->


<!-- #include file="ad_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%select case request.querystring("turn")
case "off"
db.execute("update tblCustomer set bMonitor=" & getSQLBoolean(false))
case "on"
db.execute("update tblCustomer set bMonitor=" & getSQLBoolean(true))
case "remove"
db.execute("delete from tblMonitor")
end select
if not isLeeg(request.form("sDelete")) then
db.execute("delete from tblMonitor where sDetail like '%" & left(cleanup(request.form("sDelete")),30) & "%' ")
end if%><p align="center"><span id="tot"></span></p><p align="center"><a href="ad_monitor.asp?turn=off">Turn OFF monitor for all sites</a> - 
<a href="ad_monitor.asp?turn=on">Turn ON monitor for all sites</a> - 
<a onclick="javascript:return confirm('Are you sure to remove all logs?');" href="ad_monitor.asp?turn=remove">Remove all logs</a> - 
<a href="ad_monitor.asp?show5000=on">Show 5000 logs</a></p><form action="ad_monitor.asp" method="post" name="deleteF"><table align="center"><tr><td>Delete from log where detail like </td><td><input maxlength="30" type="text" size="20" name="sDelete" /></td><td><input type="submit" name="btn" value="Go" /></td></tr></table></form><table class="sortable" align="center" border="1" cellpadding="5" cellspacing="0"><tr><td style="width:150px">Date</td><td>Details</td></tr><%dim showtot
if request.querystring("show5000")="on" then
showtot=5001
else
showtot=1001
end if
dim rs,counter,startd,endd
counter=1
set rs=db.execute("select * from tblMonitor order by dts desc")
while not rs.eof
if counter=1 then
endd=cdate(rs("dts"))
end if
if counter<showtot then
response.write "<tr><td>" & convertEuroDateTime(rs("dts")) & "</td><td><div style=""width:590px;overflow:scroll"">" & rs("sDetail") & "</div></td></tr>"
end if
startd=cdate(rs("dts"))
rs.movenext
counter=counter+1
wend
set rs=nothing%></table><%if counter>2 then
counter=counter-1%><script type="text/javascript">document.getElementById("tot").innerHTML='Started on: <%=startd%><br />Ended on: <%=endd%><br />Time elapsed: <%=datediff("s",startd,endd)%> seconds / <%=round(datediff("s",startd,endd)/3600,2)%> hours<br />Total: <%=counter%><br />Hits/second: <%=round(counter/datediff("s",startd,endd),2)%> hits/second';</script><%end if%><!-- #include file="ad_back.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
