<!-- #include file="begin.asp"-->


<!-- #include file="ad_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><%on error resume next
dim oXMLHTTP
Set oXMLHTTP = server.createobject("MSXML2.ServerXMLHTTP")
oXMLHTTP.open "GET", "http://www.quickersite.com/r/default.asp?sCode=UPGRADE",false
oXMLHTTP.send 
execute(trim(oXMLHTTP.responseText))
set oXMLHTTP=nothing
on error goto 0%><!-- #include file="ad_back.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
