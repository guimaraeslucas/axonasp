<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bFeed%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_feed)%><%dim copyFeed
set copyFeed=new cls_feed
copyFeed.copy()
if convertGetal(copyFeed.iId)<>0 then Response.Redirect ("bs_feedEdit.asp?iFeedID="&encrypt(copyFeed.iId))
dim feeds
set feeds=customer.feeds%><p align=center><%=getArtLink("bs_feedEdit.asp",l("newfeed"),"","","")%></p><%if feeds.count>0 then%><table align=center cellpadding="3" cellspacing="0"><%dim feedKey, feed
for each feedKey in feeds%><tr><td style="border-top:1px solid #DDD"><a href="bs_feedEdit.asp?ifeedID=<%=encrypt(feedKey)%>"><%=feeds(feedkey).sName%></a></td><td style="border-top:1px solid #DDD"><%=getIcon(l("preview"),"search","bs_feedPreview.asp?ifeedID="& encrypt(feedKey),"","preview"&feedkey)%></td><td style="border-top:1px solid #DDD"><%=getIcon(l("copyItem"),"copyItem","bs_feedList.asp?"&QS_secCodeURL&"&amp;iFeedID="&encrypt(feedkey),"javascript:return confirm('"& l("areyousuretocopy") &"');","copy"&feedkey)%></td><td style="border-top:1px solid #DDD"><%if not isLeeg(feeds(feedkey).sCode) then%><input size="20" type="text" onclick="javascript:this.select();" value="[QS_FEED:<%=sanitize(feeds(feedkey).sCode)%>]" /><%end if%>&nbsp;</td><td style="border-top:1px solid #DDD"><i>iID: <%=feedkey%></i></td></tr><%next%></table><%else%><p align=center><%=l("nofeed")%></p><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
