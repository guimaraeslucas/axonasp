
<%class cls_fastpage
public iId,bHideDate,sLPExternalURL,sTitle,bLPExternalOINW,sValue,iFeedId,dPage,updatedTS,sUrlRRSImage,sItemPicture,sUserfriendlyUrl,sLPIC,iPMlocation

sub class_initialize
	iPMlocation=0
end sub

function page

	set page=new cls_page
	page.iId=iId
	
	if iPMlocation<>0 then page.pick(iId)
	
end function

public function feed
set feed=new cls_feed
feed.pick(iFeedId)
end function
Public Property get sDateAndTitle
if isLeeg(dPage) then
sDateAndTitle=sTitle
else
if convertBool(bHideDate) then
sDateAndTitle=sTitle
else
sDateAndTitle=convertEuroDate(dPage)& ": " & sTitle
end if
end if
end property

	public property get listitemPicIMGTag
	
		if convertBool(customer.bListItemPic) then

			if not isLeeg(sItemPicture) then
			
				dim cStyle
				select case sLPIC
					case "fp"
						cStyle=" style=""margin:10px 0px 10px 0px;width:100%"" "
					case "al"
						cStyle=" style=""margin:5px 20px 20px 0px;width:48%;float:left"" "
					case "ar"
						cStyle=" style=""margin:5px 0px 20px 20px;width:48%;float:right"" "
					
				end select	
				
				listitemPicIMGTag="<img src=""" & C_VIRT_DIR &  Application("QS_CMS_userfiles") & "listitemimages/" & iId & "." & sItemPicture & """ " & cStyle & " class=""ListItemPictureCSS"" />"
			end if
		
		end if
		
	end property


end class%>
