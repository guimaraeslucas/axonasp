<p><%
'This is another sample script that can be added to your QuickerSite.
'This script checks if you are logged on to the backsite. 

if Session(Application("QS_CMS_iCustomerID") & "isAUTHENTICATED")=true then

	'THIS USER IS LOGGED ON THE BACKSITE! You can now assume this user is the webmaster of the site.

	'in case you want the second admin to have access also, add: or Session(Application("QS_CMS_iCustomerID") & "isAUTHENTICATEDSecondAdmin")=true

	Response.Write "You have access to the backsite. A custom module could be added here, only available to backsite users!"
	
else
	'THIS IS USER IS NOT AUTHENTICATED AS THE Webmaster.
	Response.Write "You are currently not authenticated. Please logon to the backsite of QuickerSite first!"

end if
%></p>