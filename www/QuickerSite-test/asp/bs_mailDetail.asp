<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bIntranetMail%><!-- #include file="bs_process.asp"--><%dim mail
set mail=new cls_mail
Response.Write mail.sBody
set mail=nothing%>
