
<%if page.bCanBeConvertedToFP then%><p align=center>-> <b><a onclick="javascript:return confirm('<%=l("areyousure")%>');" href="bs_convertToFP.asp?<%=QS_secCodeURL%>&amp;btnaction=ConvertToFP&amp;iId=<%=encrypt(page.iId)%>"><%=l("converttofreepage")%></a></b> <-</p><%end if%>
