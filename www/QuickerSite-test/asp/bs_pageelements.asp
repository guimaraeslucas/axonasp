<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bSetupPageElements%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Setup)%><%=getBOSetupMenu(btn_pageelements)%><table align=center style="width:450px" cellpadding="15"><tr><td style="text-align:center" colspan=4><%=l("explpageelements")%></td></tr><tr>

<td align=center>
<a style="text-decoration:none" href="bs_editbannermenu.asp">
<span class="material-symbols-outlined" style="font-size:40px">
view_column_2
</span><br /><%=l("banners")%></a>
</td>

<td align=center>
<a style="text-decoration:none" href="bs_editProps.asp">
<span class="material-symbols-outlined" style="font-size:40px">
flex_wrap
</span>
<br />Default Blocks</a>
</td>
	
<td align=center>
<a style="text-decoration:none" href="bs_editfooter.asp">
<span class="material-symbols-outlined" style="font-size:40px">
splitscreen_bottom
</span>
<br /><%=l("footer")%></a>
</td>
	
<td align=center><a style="text-decoration:none" href="bs_favicon.asp">
<span class="material-symbols-outlined" style="font-size:40px">
wb_sunny
</span>
<br /><%=l("favicon")%></a>
</td>

</tr></table><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
