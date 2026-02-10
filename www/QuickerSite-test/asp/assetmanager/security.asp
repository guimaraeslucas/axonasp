<%
Response.CacheControl		= "no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires			= -1
Response.ExpiresAbsolute	= Now()-1

if not cBool(Session(Application("QS_CMS_iCustomerID") & "isAUTHENTICATED")) and not cBool(Session(Application("QS_CMS_iCustomerID") & "isAUTHENTICATEDSecondAdmin")) and not cBool(Session(Application("QS_CMS_iCustomerID") & "isAUTHENTICATEDasUSER")) and not cBool(Session(Application("QS_CMS_iCustomerID") & "isAUTHENTICATEDIntranet")) then Response.End 
%>