<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bIntranetMail%><!-- #include file="bs_process.asp"--><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"><html><head><title><%=l("list")%></title></head><body><%dim mail, counter, receivers
set mail=new cls_mail
receivers=mail.receivers(counter)
Response.Write "<ol>"
dim i
for i=lbound(receivers,2) to ubound(receivers,2)
Response.Write "<li>" & receivers(0,i)
next
Response.Write "</ol>"
set mail=nothing%></body></html>
