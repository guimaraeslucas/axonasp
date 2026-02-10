<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bAvailabilityCal%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><!-- #include file="includes/ac_calendar.asp"--><!-- #include file="includes/ac_calendarview.asp"--><!-- #include file="includes/ac_calendarbooking.asp"--><!-- #include file="includes/ac_statuslist.asp"--><!-- #include file="bs_ac_menu.asp"--><%dim calObj,mycals,cal,calendar,cc,sCode,cv,statuslist,monthlist,monthYear,setMonth,setYear,arrD,rs,sql,i
dim action,postback,color,sDate
action=request("calAction")
postback=convertBool(request.form("postBack"))
select case lcase(action)
case "createcal","editcal"%><!-- #include file="bs_ac_calendar.asp"--><%case "embedcode"%><!-- #include file="bs_ac_embed.asp"--><%case "booking"%><!-- #include file="bs_ac_booking.asp"--><%case "bookings"%><!-- #include file="bs_ac_bookings.asp"--><%case else%><!-- #include file="bs_ac_mycals.asp"--><%end select%><%response.write message.showAlert()%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
