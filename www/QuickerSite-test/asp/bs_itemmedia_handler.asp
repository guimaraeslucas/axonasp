<!-- #include file="begin.asp"-->
<!-- #include file="bs_security.asp"-->
<!-- #include file="bs_process.asp"-->
<!-- #include file="asplite/asplite.asp"-->

<%
select case lcase(aspL.getRequest("asplEvent"))

	case "media" 	: aspl("bs_itemmedia_ajax.resx")
	
	case "upload" 	: aspl("bs_itemmedia_upload.resx")
	
end select
%>