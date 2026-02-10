<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bSetupPageElements%><!-- #include file="bs_process.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Setup)%><%=getBOSetupMenu(btn_pageelements)%><FORM method="post" action="bs_peel.asp" id=form1 name=form1> 
<input type=hidden name=btnAction value="savepeel" /><%=QS_secCodeHidden%><%if isLeeg(customer.sPeelURL) then
customer.sPeelURL="http://"
end if
if convertGetal(customer.sPeelIdleSize)=0 then
customer.sPeelIdleSize=70
end if
if convertGetal(customer.sPeelMOSize)=0 then
customer.sPeelMOSize=300
end if%><table align=center cellpadding=5><tr><td colspan=2><b>Set a Peel for your website:</b></td></tr><tr><td colspan=2><b>Step 1: Set the link to go to when clicking the peel-image:</b></td></tr><tr><td class=QSlabel>URL:*</td><td align=left><input type=text name="sPeelURL" size="50" maxlength="250" value="<%=quotrep(customer.sPeelURL)%>" /><input type=checkbox name="bPeelOINW" value="<%=true%>" <%if customer.bPeelOINW then response.write "checked"%> /> Open in new window
</td></tr><tr><td class=QSlabel>Select flip-color:*</td><td  align=left><%dim flipCounter
for  flipCounter=0 to 15
response.write "<div style=""text-align:center;float:left;margin:5px""><img style=""width:50px;height:50px""  width=""50"" height=""50"" src=""" & C_DIRECTORY_QUICKERSITE &  "/fixedImages/peels/pageflip" & flipCounter & ".png"" alt="""" />"
response.write "<br /><input type=""radio""  "
if convertGetal(customer.sPeelFlipColor)=flipCounter then
response.write " checked=""checked"" "
end if
response.write "name=""sPeelFlipColor"" value=""" & flipCounter & """ />"
response.write "</div>"
if (flipCounter+1) MOD 8=0 then response.write "<div style=""clear:both""> </div>"
next%></td></tr><tr><td class=QSlabel>Peel size (when idle):</td><td align=left><select name="sPeelIdleSize"><%=numberList(50,300,10,customer.sPeelIdleSize)%></select>px</td></tr><tr><td class=QSlabel>Peel size (when active):</td><td align=left><select name="sPeelMOSize"><%=numberList(200,700,10,customer.sPeelMOSize)%></select>px</td></tr><%if not isLeeg(customer.sPeelImage) then%><tr><td class=QSlabel>Enabled?</td><td align=left><input type=checkbox name="bPeelEnabled" value="<%=true%>" <%if customer.bPeelEnabled then response.write "checked"%> /></td></tr><%end if%><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td>&nbsp;</td><td><INPUT class="art-button" TYPE=SUBMIT VALUE="<%=l("save")%>" id=SUBMIT1 name=SUBMIT1 /> <%if not isLeeg(customer.sPeelImage) then%><INPUT class="art-button" TYPE=SUBMIT onclick="javascript:return confirm('<%=l("areyousure")%>');" VALUE="<%=l("delete")%>" id=SUBMIT1 name=btnDelete /><%end if%></td></tr><%if customer.sPeelURL<>"http://" then%><tr><td colspan=2><b>Step 2: <a href="bs_selectPeel.asp">Upload your peel-image (PNG/JPG/GIF file)</a></b></td></tr><%if not isLeeg(customer.sPeelImage) then%><tr><td colspan=2>Current peel image: <a target="_blank" href="<%=C_VIRT_DIR &  Application("QS_CMS_userfiles") & customer.sPeelImage%>">(download)</a><br /><div style="width:307px;height:308px;text-align:right;margin:20px;background:url(<%select case right(customer.sPeelImage,3) 
case "jpg"
response.write C_DIRECTORY_QUICKERSITE & "/showthumb.aspx?maxsize=308&FSR=1&img=" & C_VIRT_DIR &  Application("QS_CMS_userfiles") & customer.sPeelImage & ") no-repeat right top;"
case else
response.write C_VIRT_DIR &  Application("QS_CMS_userfiles") & customer.sPeelImage & ") no-repeat right top;"
end select%>" /><img alt="" src="<%=C_DIRECTORY_QUICKERSITE%>/fixedImages/peels/pageflip<%=customer.sPeelFlipColor%>.png" /></div></td></tr><%end if%><%end if%></table></form><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
