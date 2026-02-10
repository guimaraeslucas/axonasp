<!-- #include file="asp/begin.asp"--><%printReplies=true%><!-- #include file="asp/process.asp"--><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" dir="<%=tdir%>" lang="en-US" xml:lang="en">
<head>
<title><%=selectedPage.showTitle%></title>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=<%=QS_CHARSET%>" />
<link rel="stylesheet" TYPE="text/css" href="<%=C_DIRECTORY_QUICKERSITE%>/css/qs_<%=tdir%>.css" title="QSStyle" />
<script type="text/javascript" src="<%=C_DIRECTORY_QUICKERSITE%>/js/sorttable.js"></script>
<%if not isLeeg(selectedPage.sApplication) then Response.End%>
</head>
<body onload="javascript: window.focus();window.print();">
<b><%=treatConstants(selectedPage.replaceBlocks(pageTitle),true)%></b><hr /><%=treatConstants(selectedPage.replaceBlocks(pageBody),true)%>
</body>
</html>
<%cleanUPASP%>
