<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bSetupPageElements%><%dim page
set page=new cls_page
page.pick(decrypt(request("iPageID")))
if request("btnaction")=l("save") then
checkCSRF()
page.sProp01=removeEmptyP(convertStr(Request.Form ("sProp01")))
page.sProp02=removeEmptyP(convertStr(Request.Form ("sProp02")))
page.sProp03=removeEmptyP(convertStr(Request.Form ("sProp03")))
page.sProp04=removeEmptyP(convertStr(Request.Form ("sProp04")))
page.sProp05=removeEmptyP(convertStr(Request.Form ("sProp05")))
page.sProp06=removeEmptyP(convertStr(Request.Form ("sProp06")))
page.sProp07=removeEmptyP(convertStr(Request.Form ("sProp07")))
page.sProp08=removeEmptyP(convertStr(Request.Form ("sProp08")))
if page.save then message.Add ("fb_saveOK")
end if%><!-- #include file="includes/commonheader.asp"--><body style="background-color:#FFF"><!-- #include file="bs_initBack.asp"--><p align="center">Below you can provide content for the <strong>8 default blocks</strong> for the page: <strong><%=page.sTitle%></strong>.</p><form action="bs_editPageBlocks.asp" method="post" name=mainform><%=QS_secCodeHidden%><input type="hidden" value="<% =l("save")%>" name=btnaction /><input type="hidden" value="<% =encrypt(page.iId)%>" name=iPageID /><table style="width:700px" align=center cellpadding=4 cellspacing=0><tr><td>[PAGE_BLOCK01]</td><td><textarea style="width:480px" cols="30" rows="3" name=sProp01><%=quotrep(page.sProp01)%></textarea></td></tr><tr><td align="right"><i>Default value:</i></td><td><%=customer.sProp01%></td></tr><tr><td colspan="2"><hr/></td></tr><tr><td>[PAGE_BLOCK02]</td><td><%createFCKInstance page.sProp02,"siteBuilderMailSource","sProp02"%></td></tr><tr><td align="right"><i>Default value:</i></td><td><%=customer.sProp02%></td></tr><tr><td colspan="2"><hr/></td></tr><tr><td>[PAGE_BLOCK03]</td><td><textarea style="width:480px" cols="30" rows="3" name=sProp03><%=quotrep(page.sProp03)%></textarea></td></tr><tr><td align="right"><i>Default value:</i></td><td><%=customer.sProp03%></td></tr><tr><td colspan="2"><hr/></td></tr><tr><td>[PAGE_BLOCK04]</td><td><%createFCKInstance page.sProp04,"siteBuilderMailSource","sProp04"%></td></tr><tr><td align="right"><i>Default value:</i></td><td><%=customer.sProp04%></td></tr><tr><td colspan="2"><hr/></td></tr><tr><td>[PAGE_BLOCK05]</td><td><textarea style="width:480px" cols="30" rows="3" name=sProp05><%=quotrep(page.sProp05)%></textarea></td></tr><tr><td align="right"><i>Default value:</i></td><td><%=customer.sProp05%></td></tr><tr><td colspan="2"><hr/></td></tr><tr><td>[PAGE_BLOCK06]</td><td><%createFCKInstance page.sProp06,"siteBuilderMailSource","sProp06"%></td></tr><tr><td align="right"><i>Default value:</i></td><td><%=customer.sProp06%></td></tr><tr><td colspan="2"><hr/></td></tr><tr><td>[PAGE_BLOCK07]</td><td><textarea style="width:480px" cols="30" rows="3" name=sProp07><%=quotrep(page.sProp07)%></textarea></td></tr><tr><td align="right"><i>Default value:</i></td><td><%=customer.sProp07%></td></tr><tr><td colspan="2"><hr/></td></tr><tr><td>[PAGE_BLOCK08]</td><td><%createFCKInstance page.sProp08,"siteBuilderMailSource","sProp08"%></td></tr><tr><td align="right"><i>Default value:</i></td><td><%=customer.sProp08%></td></tr><tr><td colspan="2"><hr/></td></tr><tr><td>&nbsp;</td><td><input class="art-button" type="submit" value="<%=l("save")%>" name="dummy" /></td></tr></table></form><!-- #include file="bs_endBack.asp"--><%=message.showAlert()%><%=cPopup.dumpJS()%> 
</body></html><%cleanUPASP%>
