<!-- #include file="begin.asp"-->


<!-- #include file="ad_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><!-- #include file="iismanager/includes.asp"--><%on error resume next
dim webobj,i
set webobj = GetObject("IIS://localhost/w3svc/" & request("name"))
if webobj.name="" then response.end
if err.number<>0 then response.end
if request.form("btnSaveIISB")="Save" then
dim sbArr,sbKey,sbDict
set sbDict=server.createObject("scripting.dictionary")
sbArr=split(request.form("bindings"),vbcrlf)
for sbKey=lbound(sbArr) to ubound(sbArr)
if not sbDict.exists(sbArr(sbKey)) then
if not isLeeg(sbArr(sbKey)) then
sbDict.add replace(lcase(trim(sbArr(sbKey))),"http://","",1,-1,1),""
end if
end if
next
if sbDict.count>0 then
dim cc
cc=sbDict.count-1
dim newSBS(),iSBS
redim newSBS(cc)
iSBS=0
for each sbKey in sbDict
newSBS(iSBS)=":80:" & sbKey
iSBS=iSBS+1
next
webobj.ServerBindings=newSBS
webobj.SetInfo()
dim bs
for i=lbound(webobj.serverbindings) to ubound(webobj.serverbindings)
bs=bs& webobj.serverbindings(i)(i) & vbcrlf
next
dim rs
set rs=GetDynamicRS
rs.open "select * from sites where Name='" & webobj.name & "'"
if not rs.eof then
rs("ServerBindings")=bs
rs.update()
end if
set rs=nothing
response.redirect ("ad_editBindings.asp?name=" & webobj.name & "&fbMessage=fb_saveOK")
end if
response.redirect ("ad_editBindings.asp?name=" & webobj.name & "&strMessage=err_mandatory")
end if%><p>Edit bindings for <b><%=webobj.serverComment%></b>:</p><form method="post" action="ad_editbindings.asp" name="mainform"><input type="hidden" name="name" value="<%=quotrep(request("name"))%>" /><table align="center" cellpadding="5" border="0"><tr><td><font color=Red><b>ENTER-Separated list of domains - DO NOT INCLUDE "http://"</b></font></td></tr><tr><td><textarea name=bindings cols="90" rows="15"><%for i=lbound(webobj.serverbindings) to ubound(webobj.serverbindings)
response.write replace(webobj.serverbindings(i)(i),":80:","",1,-1,1) & vbcrlf
next%></textarea></td></tr><tr><td><input class="art-button" type="submit" name="btnSaveIISB" value="Save" /></td></tr></form><table align=center><tr><td align=center>-> <b><a href="ad_iis.asp">Back</a></b> <-</td></tr></table><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
