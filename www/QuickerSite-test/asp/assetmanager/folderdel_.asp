<!-- #include file="security.asp"-->
<%
If Not request("inpCurrFolder")="" Then
	sFolder = request("inpCurrFolder")
		
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(sFolder) Then
		sDestination = fso.GetParentFolderName(sFolder)

		fso.DeleteFolder(sFolder)
		sMsq = "<script>document.write(getTxt('Folder deleted.'))</script>"
	Else
		sMsq = "<script>document.write(getTxt('Folder does not exist.'))</script>"
	End If
	Set fso = Nothing
End If
%>
<base target="_self">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="style.css" rel="stylesheet" type="text/css">
<script>
	if(navigator.appName.indexOf('Microsoft')!=-1)
		var sLang=dialogArguments.sLang;
	else
		var sLang=window.opener.sLang;
	document.write("<scr"+"ipt src='language/"+sLang+"/folderdel_.js'></scr"+"ipt>");
</script>
<script>writeTitle()</script>
<script>
function refresh()
	{
	if(navigator.appName.indexOf('Microsoft')!=-1)
		dialogArguments.refreshAfterDelete(inpDest.value);
	else
		window.opener.refreshAfterDelete(document.getElementById("inpDest").value);
	}
</script>
</head>
<body onload="loadTxt()" style="overflow:hidden;margin:0px;">

<table width=100% height=100% align=center style="" cellpadding=0 cellspacing=0 ID="Table1">
<tr>
<td valign=top style="padding-top:5px;padding-left:15px;padding-right:15px;padding-bottom:12px;height=100%">

	<br />
	<input type="hidden" ID="inpDest" NAME="inpDest" value="<%=sDestination%>">
	<div><b><%=sMsq%>&nbsp;</b></div>

</td>
</tr>
<tr>
<td class="dialogFooter" style="height:45px;padding-right:10px;" align=right valign=middle>
	<input type="button" name="btnCloseAndRefresh" id="btnCloseAndRefresh" value="close & refresh" onclick="refresh();self.close();" class="inpBtn" onmouseover="this.className='inpBtnOver';" onmouseout="this.className='inpBtnOut'">
</td>
</tr>
</table>


</body>
</html>