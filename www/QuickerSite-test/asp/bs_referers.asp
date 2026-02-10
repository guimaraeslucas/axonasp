<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bStats%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Setup)%><%=getBOSetupMenu(btn_Stats)%><%Function URLDecode2(str) 
str=convertStr(str)
dim i,sT,sR
on error resume next
    str = Replace(str, "+", " ",1,-1,1) 
    For i = 1 To Len(str) 
        sT = Mid(str, i, 1) 
        If sT = "%" Then 
            If i+2 < Len(str) Then 
                sR = sR & _ 
                    Chr(CLng("&H" & Mid(str, i+1, 2))) 
                i = i+2 
            End If 
        Else 
            sR = sR & sT 
        End If 
    Next 
    URLDecode2 = sR 
err.clear
on error goto 0
End Function %><%if not customer.bScanreferer then Response.Redirect ("bs_stats")
dim dFrom,dUntill,groupBy
groupBy	= Request.Form ("groupBy")
dUntill	= Request.Form ("dUntill")
dFrom	= Request.Form ("dFrom")
if dFrom="" then dFrom=convertEuroDate(dateAdd("d",-7,date()))
if dUntill="" then dUntill=convertEuroDate(date())
if groupBy="" then groupBy=sb_ref_REFERER
dim refGroupByList,re
set refGroupByList=new cls_refGroupByList%><form method="post" action="bs_referers.asp" name="mainform"><input type="hidden" name="postback" value="<%=true%>" /><table align="center" cellpadding="2"><tr><td><%=l("groupby")%>:</td><td><select onchange="javascript:document.mainform.submit();" id=select1 name=groupBy><%=refGroupByList.showSelected("option",groupBy)%></select></td><td><%=l("from")%>:</td><td><input type="text" id="dFrom" name="dFrom" value="<%=dFrom%>" /><%=JQDatePicker("dFrom")%></td><td><%=l("untill")%>:</td><td><input type="text" id="dUntill" name="dUntill" value="<%=dUntill%>" /><%=JQDatePicker("dUntill")%></td><td><input class="art-button" type=submit value="Go!" name="btnaction" /></td></tr></table></form><%if Request.QueryString ("pageAction")="Reset" then
checkCSRF()
removeReferers()
end if
dim arrRef, total, ref, refDict,keyValue,ri,itemValue
arrRef=getReferersArr(dFrom,dUntill,total)
if isArray(arrRef) then
set refDict=server.CreateObject ("scripting.dictionary")
if sb_ref_QUERY<>groupBy then%><p align=center><%=l("total")%>:&nbsp;<b><%=total%></b></p><%end if
select case groupBy
case sb_ref_REFERER,sb_ref_URL,sb_ref_USERIP,sb_ref_LP,sb_ref_LA,sb_ref_QUERY%><table cellpadding=4 cellspacing=0 align=center class=sortable id=ref><tr><th width=450><%=refGroupByList.showSelected("single",groupBy)%></th><th><%=l("number")%></th></tr><%select case groupBy
case sb_ref_REFERER
keyValue=0
case sb_ref_URL
keyValue=1
case sb_ref_LP
keyValue=2
case sb_ref_USERIP
keyValue=3
case sb_ref_LA
keyValue=4
case sb_ref_QUERY
keyvalue=999
end select
for ref=lbound(arrRef,2) to ubound(arrRef,2)
select case keyValue
 case 0 ' referer URL
itemValue="<a href=""" & quotrep(arrRef(keyValue,ref)) & """ target='eP'>" & quotrep(left(arrRef(keyValue,ref),80)) & "</a>"
 case 1 ' referrer site
itemValue="<a href=""" & quotrep(filterURL(arrRef(0,ref))) & """ target='eP'>" & quotrep(filterURL(arrRef(0,ref))) & "</a>"	 
 case 3 ' userIP
itemValue=quotrep(arrRef(keyValue,ref))
 case 2 ' landingpage
itemValue="<a href=""" & quotrep(arrRef(keyValue,ref)) & """ target='eP'>" & quotrep(arrRef(keyValue,ref)) & "</a>"
 case 4 ' language
itemValue=quotrep(left(arrRef(keyValue,ref),5))
 case 999 'zoekterm
itemValue=quotrep(trim(lcase(URLDecode2(filterQuery(arrRef(0,ref))))))
end select
if itemValue<>"" then
if not refDict.exists(itemValue) then
refDict.Add itemValue,1
else
refDict(itemValue)=convertGetal(refDict(itemValue)+1)
end if
end if
next
for each ri in refDict%><tr><td width=450 style="border-top:1px solid #DDD"><%=ri%></td><td align=center style="border-top:1px solid #DDD"><%=refDict(ri)%></td></tr><%next%></table><%case sb_ref_LOGFILE%><table cellpadding=2 align=center class=sortable id=ref><tr><th width=350><%=l("referringsites")%></th><th><%=l("landingpage")%></th><th><%=l("language")%></th><th>IP</th><th><%=l("date")%></th></tr><%for ref=lbound(arrRef,2) to ubound(arrRef,2)%><tr><td width=350 valign=top><a href="<%=quotrep(arrRef(0,ref))%>" target="ref<%=quotrep(ref)%>"><%=quotrep(urlDecode(arrRef(0,ref)))%></a></td><td valign=top><a target="ref<%=quotrep(ref)%>" href="<%=quotrep(arrRef(2,ref))%>"><%=quotrep(arrRef(2,ref))%></a></td><td align=center valign=top><%=quotrep(left(arrRef(4,ref),5))%></td><td align=center valign=top><%=quotrep(arrRef(3,ref))%></td><td align=center valign=top><%=formatTimeStamp(arrRef(5,ref))%></td></tr><%next%></table><%end select
set refDict=nothing
end if
set refGroupByList=nothing%><p align=center><a onclick="javascript:return confirm('<%=l("areyousure")%>');" href="bs_referers.asp?<%=QS_secCodeURL%>&amp;pageAction=Reset"><b><%=l("reset")%></b></a>&nbsp;-&nbsp;<b><a href="bs_stats.asp"><%=l("stats")%></a></b></p><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
