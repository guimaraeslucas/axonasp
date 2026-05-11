<%@ Language=VBScript %>
<%
Option Explicit

Dim a, b, c
Dim sinv, cosv, tanv, atnv, sqrv, absv, expv, logv, roundv, intv

a = 9
b = 3
c = a + b * 2 - 1

sinv = Sin(0)
cosv = Cos(0)
tanv = Tan(0)
atnv = Atn(1)
sqrv = Sqr(81)
absv = Abs(-12)
expv = Exp(1)
logv = Log(1)
roundv = Round(2.5)
intv = Int(3.9)

Response.Write "INT_ARITH=" & c & vbCrLf
Response.Write "SHIFT_DIV=" & (64 \ 4) & vbCrLf
Response.Write "MATH=" & sinv & "|" & cosv & "|" & tanv & "|" & atnv & "|" & sqrv & "|" & absv & "|" & expv & "|" & logv & "|" & roundv & "|" & intv
%>