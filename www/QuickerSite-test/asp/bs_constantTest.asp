<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bHomeVBScript%><!-- #include file="includes/commonheader.asp"--><body style="background-color:#FFF"><!-- #include file="bs_initBack.asp"--><%if Request.Form ("sParameters")<>"" then%><form action="bs_constantTest.asp" name="tC" method="post"><%=QS_secCodeHidden%>Parameter(s): <b><%=sanitize(Request.Form ("sParameters"))%></b><br /><i>Include values below <b>(between double quotes!)</b></i><br /><input type=hidden name=sParameters value="<%=sanitize(Request.Form ("sParameters"))%>" /><input type=hidden name=customScript value="<%=sanitize(Request.Form ("customScript"))%>" /><input type=text value="<%=sanitize(Request.Form ("sParametersValues"))%>" name=sParametersValues /><input type=text value="<%=sanitize(Request.Form ("sGlobal"))%>" name=sGlobal /><input class="art-button" type=submit value="Test!" /></form><%end if
checkCSRF()
if isLeeg(Request.Form ("sParameters")) or (Request.Form ("sParameters")<>"" and not isLeeg(Request.Form ("sParametersValues"))) then
Response.Write executeConstant(Request.Form ("customScript")&QS_VBScriptIdentifier&Request.Form ("sParameters"),true,Request.Form ("sParametersValues"),Request.Form ("sGlobal"))
end if%><!-- #include file="bs_endBack.asp"--><%=message.showAlert()%><%=cPopup.dumpJS()%> 
</body></html><%cleanUPASP%>
