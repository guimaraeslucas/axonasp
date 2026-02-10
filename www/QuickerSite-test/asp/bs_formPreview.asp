<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bForms%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml" dir="<%=tdir%>" lang="en-US" xml:lang="en"><head><title>Form</title><meta HTTP-EQUIV="Content-Type" content="text/html; charset=<%=QS_CHARSET%>" /><link rel="stylesheet" TYPE="text/css" href="<%=C_DIRECTORY_QUICKERSITE%>/css/qs_<%=tdir%>.css" title="QSStyle" /><script type="text/javascript" src="<%=C_DIRECTORY_QUICKERSITE%>/js/sorttable.js"></script><%=dumpJavaScript(true)%><%=css()%></head><body style="background-color:#FFFFFF"><table style="width:680px"><tr><td><%dim form
set form=new cls_form
Response.write form.build("bs_formPreview.asp","center","button",null)%></td></tr></table></body></html><%cleanUPASP%>
