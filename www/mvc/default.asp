<%
@Language = "VBSCRIPT" CodePage = "65001"
%>
<%
Option Explicit
%>
<!--#include file="model/site_model.asp" -->
<!--#include file="controller/site_controller.asp" -->
<%
Dim route
Dim htmlOutput

route = LCase(Trim(Request.QueryString("page")))

If route = "" Then
    route = "home"
End If

Select Case route
    Case "home"
        htmlOutput = RenderHomePage()
    Case "example"
        htmlOutput = RenderExamplePage()
    Case Else
        htmlOutput = RenderHomePage()
End Select

Response.Write htmlOutput
%>
