<!-- #include file="begin.asp"-->
<!-- #include file="ad_security.asp"-->
<!-- #include file="includes/header.asp"-->
<!-- #include file="bs_initBack.asp"-->
<!-- #include file="bs_header.asp"-->
<%dim customerList,getAll
set customerList=new cls_customerList
set getALL=customerList.getAll%><table style="border-style:none" align="center" cellpadding="4" cellspacing="0"><tr><td>Date</td><td>Message</td><td>Site</td></tr><%dim rs, sql, rs2
sql="SELECT top 50 * from tblGuestBookItem order by dCreatedTS desc"
set rs=db.execute(sql)
while not rs.eof
'response.write "select tblCustomer.sName,tblCustomer.sUrl from tblCustomer join tblGuestbook on tblGuestbook.iCustomerID=tblCustomer.iId where tblGuestbook.iId=" & rs("iGuestBookID")
'response.end 
set rs2=db.execute("select tblCustomer.sName,tblCustomer.sUrl from tblCustomer inner join tblGuestbook on tblGuestbook.iCustomerID=tblCustomer.iId where tblGuestbook.iId=" & convertGetal(rs("iGuestBookID")))
if not rs2.eof then
response.write "<tr>"
response.write "<td style=""border-bottom:1px solid #DDD"">" & convertEurodate(rs("dCreatedTS")) & "</td>"
response.write "<td style=""border-bottom:1px solid #DDD;max-width: 560px;   overflow: hidden;   white-space: wrap;"" >" & left(quotrep(rs("sValue")),400) & "...</td>"
response.write "<td style=""border-bottom:1px solid #DDD""><a target=""_new"" href=""" & rs2("surl") & """>"  & rs2("sName") & "</a></td>"
response.write "</tr>"
end if
set rs2=nothing
rs.movenext
wend
set rs=nothing%></table>
<!-- #include file="ad_back.asp"-->
<!-- #include file="bs_endBack.asp"-->
<!-- #include file="includes/footer.asp"-->
