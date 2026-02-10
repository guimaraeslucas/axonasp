
<%set calendar=new cls_calendar
calendar.pick(request("iCalID"))
if postback then
select case request.form("btnName")
case "Save"
calendar.sName=request.form("sName")
if calendar.save() then
response.redirect("bs_ac.asp")
end if
case "Delete"
calendar.delete()
response.redirect("bs_ac.asp")
end select
end if%><form action="bs_ac.asp" method="post"><input type="hidden" value="<%=calendar.iId%>" name="iCalID" /><input type="hidden" value="editCal" name="calAction" /><input type="hidden" value="<%=true%>" name="postback" /><table><tr><td class=QSlabel>Name of Calendar:*</td><td><input type="text" size="30" maxlength="50" name="sName" value="<%=quotrep(calendar.sName)%>" /><small> (in most cases the name of the property)</small></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type="submit" name="btnName" value="Save" /> <%if convertGetal(calendar.iId)<>0 then%><input  class="art-button" onclick="javascript:return confirm('Are you sure to delete this calendar? NO WAY TO UNDO!')" type="submit" name="btnName" value="Delete" /><%end if%></td></tr></table></form><%set calendar=nothing%>
