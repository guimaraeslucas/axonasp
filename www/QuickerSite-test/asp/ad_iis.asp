<!-- #include file="begin.asp"-->


<!-- #include file="ad_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><!-- #include file="iismanager/includes.asp"--><%dim rs,sql,sValue
sql="select * from sites"
sValue=left(replace(request.form("sValue"),"'","''",1,-1,1),150)
if not isLeeg(sValue) then
sql=sql &" where 1=1 and "
sql=sql &"("
sql=sql & "serverbindings like '%" & sValue & "%' OR "
sql=sql & "path like '%" & sValue & "%' OR "
sql=sql & "servercomment like '%" & sValue & "%' OR "
sql=sql & "QSVD like '%" & sValue & "%' OR "
sql=sql & "ServerState like '%" & sValue & "%' OR "
sql=sql & "QSPath like '%" & sValue & "%'"
sql=sql & ")"
end if
sql=sql&" order by servercomment asc"
set rs=getConn.execute(sql)
if request.querystring("reload")="1" then 
loadIIS()
response.redirect("ad_iis.asp")
end if%><form method=post action=ad_iis.asp name=mainform><p align=center><span class="art-button-wrapper"><span class="art-button-l"> </span><span class="art-button-r"> </span><a onclick="javascript:return confirm('Are you sure to reload IIS? This can take a while.');" class="art-button" href="ad_iis.asp?reload=1">Reload IIS</a></span>  
<input type=text style="margin-left:10px;margin-right:10px" size="30" value="<%=quotrep(request.form("sValue"))%>" name="sValue" /> <span class="art-button-wrapper"><span class="art-button-l"> </span><span class="art-button-r"> </span><input class="art-button" type="submit" name="btnSearch" value="Search" /></span></p></form><p align="center"><span id="tot">Please wait...</span> <span id="totsize"></span></p><table align="center" cellpadding="3" cellspacing="0" border=1 style="width:95%"><tr><td><b>Comments</b></td><td><b>Bindings</b></td></tr><%dim cccc,totsize,dictBindings
set dictbindings=server.createobject("scripting.dictionary")
cccc=0
totsize=0
while not rs.eof
cccc=cccc+1%><tr><td valign="top"><b><%=rs("servercomment")%></b><ul><li><b>Path:</b> <%=rs("path")%></li><li><b>State:</b> <%=rs("ServerState")%></li><li><b>QS Path:</b> <%=rs("QSPath")%></li><li><b>QS Dir:</b> <%=rs("QSVD")%></li></ul></td><td  valign="top"><ul><%dim arrB,b
arrB=split(rs("serverbindings"),":80:")
for b=lbound(arrB) to ubound(arrB)
totsize=totsize+1%><li><a target="_blank" href="http://<%=arrB(b)%>"><%=arrB(b)%></a></li><%if not isLeeg(arrB(b)) then
if not dictbindings.exists(arrB(b)) then
dictbindings.Add arrB(b),""
end if
end if
next
totsize=totsize-1%><li style="background-image:none"><span style="margin-top:6px" class="art-button-wrapper"><span class="art-button-l"> </span><span class="art-button-r"> </span><a class="art-button" href="ad_editbindings.asp?name=<%=rs("name")%>">Modify Bindings</a></span></li></ul></td></tr><%rs.movenext
wend
set rs=nothing%></table><p align=center><b>List of all bindings:</b><textarea style="width:500px;height:400px"><%for each b in dictbindings
response.write b
next%></textarea></p><script type="text/javascript">document.mainform.sValue.focus();</script><script type="text/javascript">document.getElementById("tot").innerHTML='<b><%=cccc%></b> sites /'</script><script type="text/javascript">document.getElementById("totsize").innerHTML='<b><%=totsize%></b> bindings'</script><!-- #include file="ad_back.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
