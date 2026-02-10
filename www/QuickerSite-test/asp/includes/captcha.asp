<%@LANGUAGE="VBSCRIPT"%>
<%
'ASP Security Image Generator v3.0 - 13/September/2006
'Generate images to make a CAPTCHA test
'© 2006 Emir Tüzül. All rights reserved.
'http://www.tipstricks.org

'This program is free software; you can redistribute it and/or
'modify it under the terms of the Common Public License
'as published by the Open Source Initiative OSI; either version 1.0
'of the License, or any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
'Common Public License for more details.

'http://www.tipstricks.org/aspsig/fontmap.asp
'*[null pixel]Numbers[repeat count], #[text]Numbers[repeat count], &[row reference]number[referenced row index]
'First row [font height, chars...]
'Following rows [char width, pixel maps...]
FontMap = Array(_
split("13,A,B,C,D,E,F,0,1,2,3,4,5,6,7,8,9",",") ,_
split("14,*5#4*5,*4#6*4,&2,&2,*3#3*2#3*3,&5,*2#4*2#4*2,*2#3*4#3*2,*2#10*2,*1#12*1,*1#3*6#3*1,&11,#3*8#3",",") ,_
split("11,#8*3,#10*1,#3*4#3*1,&3,&3,&1,&2,#3*4#4,#3*5#3,&9,&8,&2,#9*2",",") ,_
split("11,*4#6*1,*2#9,*1#4*4#2,*1#3*6#1,#3*8,&5,&5,&5,&5,&4,&3,&2,&1",",") ,_
split("12,#8*4,#10*2,#3*4#4*1,#3*5#3*1,#3*6#3,&5,&5,&5,&5,&4,&3,&2,&1",",") ,_
split("9,#9,&1,#3*6,&3,&3,#8*1,&6,&3,&3,&3,&3,&1,&1",",") ,_
split("9,#9,&1,#3*6,&3,&3,&1,&1,&3,&3,&3,&3,&3,&3",",") ,_
split("11,*3#5*3,*1#9*1,*1#3*3#3*1,#3*5#3,&4,&4,&4,&4,&4,&4,&3,&2,&1",",") ,_
split("9,*3#3*3,&1,#6*3,&3,*3#3*3,&5,&5,&5,&5,&5,&5,#9,&12",",") ,_
split("10,*1#6*3,#8*2,#2*3#4*1,#1*5#3*1,*6#3*1,&5,*5#3*2,*4#4*2,*3#4*3,*2#4*4,*1#4*5,#10,&12",",") ,_
split("11,*1#8*2,#10*1,#3*5#3,#1*7#3,*7#3*1,*3#6*2,*3#7*1,*7#4,*8#3,&4,#3*4#4,&2,*1#7*3",",") ,_
split("12,*6#4*2,*5#5*2,&2,*4#2*1#3*2,*3#3*1#3*2,*2#3*2#3*2,*1#3*3#3*2,#3*4#3*2,#12,&9,*7#3*2,&11,&11",",") ,_
split("11,*1#10,&1,*1#3*7,&3,*1#8*2,*1#9*1,*7#4,*8#3,&8,#1*7#3,#3*4#3*1,#10*1,*1#7*3",",") ,_
split("11,*4#6*1,*2#8*1,*1#4*6,*1#3*7,#3*1#5*2,#10*1,#3*4#4,#3*5#3,&8,&8,*1#3*3#3*1,*1#9*1,*3#5*3",",") ,_
split("11,#11,&1,*7#4,*7#3*1,*6#4*1,*6#3*2,*5#3*3,*4#4*3,*4#3*4,*3#4*4,*3#3*5,*2#3*6,*1#4*6",",") ,_
split("11,*2#7*2,*1#9*1,#3*4#4,#3*5#3,#4*3#3*1,*1#8*2,&1,*1#3*1#5*1,&4,&4,#4*3#4,&2,*2#6*3",",") ,_
split("11,*3#5*3,*1#9*1,*1#3*3#3*1,#3*5#3,&4,&4,#4*4#3,*1#10,*2#5*1#3,*7#3*1,*6#4*1,*1#8*2,*1#6*4",",") _
)'Previous row must end with _

'#Begin ColorMap
'http://www.tipstricks.org/aspsig/colormap.asp
const BmpColorMap = "dffeff000c851700c0c0c0004d74de00e9dad100c9634d009cc9d600633d1f009600000078fcf500e1db92003132ac004763fe0033ffad00bcb9f10043480b00eceeee006c363600"

ColorMap = Array(_
split("00,01,01",",") ,_
split("02,03,03",",") ,_
split("04,05,05",",") ,_
split("06,07,07",",") ,_
split("08,09,09",",") ,_
split("0A,0B,0B",",") ,_
split("0C,0D,0D",",") ,_
split("00,05,05",",") ,_
split("0E,0F,0F",",") ,_
split("10,11,11",",") _
)'End ColorMap

'#Auto calculated variables
dim ImageWidth, ImageHeight, arrTextWidth(), TextHeight, LeftMargin, arrTopMargin(), CursorPos
dim BmpEndLine, BColor, TColor, NColor
dim i, j, k, x, y

'#Editable consts and variables
dim Bitmap(25,62) '[Height,Width]
const CodeLength = 4 'Secure code length (Max:8)
const CharTracking = 3 'Set the tracking between two characters
const RndTopMargin = false 'Randomize top margin every character
const NoiseEffect = 0 '0[none], 1[sketch], 2[random foreground lines], 3[random background lines], 4[1 and 3 (Recommed maximum NoiseLine=4)]
const NoiseLine = 2 'Low values make easy OCR, high values decrease readability
const MinLineLength = 6 'Minimum noise line length
'const MaxLineLength = 8 'Maximum noise line length (Not available at this time)

'#Subroutines and functions
function CreateGUID(valLength)
	const strValid = "ABCDEF1234567890"
	tmpGUID = vbNullString
	tmpChr = vbNullString
	Randomize(Timer)
	for cGUID=1 to valLength
		do 
			tmpChr = Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
		loop while CStr(tmpChr) = CStr(Right(tmpGUID,1))
		tmpGUID = tmpGUID & tmpChr
	Next
	CreateGUID = tmpGUID
end function

function RndInterval(valMin,valMax)
	Randomize(Timer)
	RndInterval = Int(((valMax - valMin + 1) * Rnd()) + valMin)
end function

function GetCharMap(valChr)
	dim i, j
	j = 0
	for i=1 to UBound(FontMap(0))
		if CStr(FontMap(0)(i)) = CStr(valChr) then
			j = i
			exit for
		end if
	next

	if j > 0 then
		GetCharMap = FontMap(j)
	else
		GetCharMap = Array(0)
	end if
end function

sub WriteCanvas(byval valChr, byval valTopMargin)
	dim i, j, k, curPos, tmpChr, arrChrMap, strPixMap, drawPixel, pixRepeat

	'find char map
	arrChrMap = GetCharMap(valChr)
	if UBound(arrChrMap) < 1 then
		exit sub
	end if

	'write char
	for i=1 to UBound(arrChrMap)
		'get pixel map active line
		strPixMap = arrChrMap(i)
		if Left(strPixMap,1) = "&" then
			j = Mid(strPixMap,2)
			if (IsNumeric(j) = true) then
				strPixMap = arrChrMap(CInt(j))
			else
				strPixMap = vbNullString
			end if
		end if
		strPixMap = Trim(strPixMap)

		'drawing pixel
		curPos = CursorPos
		drawPixel = false
		pixRepeat = vbNullString
		for j=1 to Len(strPixMap)
			tmpChr = Mid(strPixMap,j,1)
			if (IsNumeric(tmpChr) = true) and (j < Len(strPixMap)) then
				pixRepeat = pixRepeat & tmpChr
			else
				'end pixel map?
				if IsNumeric(tmpChr) = true then
					pixRepeat = pixRepeat & tmpChr
				end if

				'draw pixel
				if (drawPixel = true) and (IsNumeric(pixRepeat) = true) then
					for k=1 to CInt(pixRepeat)
						curPos = curPos + 1
						Bitmap((valTopMargin + i),curPos) = TColor
					next
				elseif IsNumeric(pixRepeat) = true then
					curPos = curPos + CInt(pixRepeat)
				end if

				'what is new command?
				if tmpChr = "#" then
					drawPixel = true
				else
					drawPixel = false
				end if
				pixRepeat = vbNullString
			end if
		next
	next
end sub

sub PrepareBitmap(valSecureCode)
	dim i, j
	'image dimensions
	ImageWidth = UBound(Bitmap,2)
	ImageHeight = UBound(Bitmap,1)

	'char and text width
	redim arrTextWidth(CodeLength)
	arrTextWidth(0) = 0
	for i=1 to CodeLength
		arrTextWidth(i) = CInt(GetCharMap(Mid(secureCode,i,1))(0))
		arrTextWidth(0) = arrTextWidth(0) + arrTextWidth(i)
	next
	arrTextWidth(0) = arrTextWidth(0) + ((CodeLength - 1) * CharTracking)

	'text height
	TextHeight = CInt(FontMap(0)(0))

	'left margin
	LeftMargin = Round((ImageWidth - arrTextWidth(0)) / 2)

	'top margin
	redim arrTopMargin(CodeLength)
	arrTopMargin(0) = Round((ImageHeight - TextHeight) / 2)
	if RndTopMargin = true then
		for i=1 to CodeLength
			arrTopMargin(i) = RndInterval(Int(arrTopMargin(0) / 2),(arrTopMargin(0) + Round(arrTopMargin(0) / 2)))
		next
	else
		for i=1 to CodeLength
			arrTopMargin(i) = arrTopMargin(0)
		next
	end if

	'color selection
	i = RndInterval(0,UBound(ColorMap))
	BColor = ColorMap(i)(0)
	NColor = ColorMap(i)(1)
	TColor = ColorMap(i)(2)

	'Apply background effect
	if NoiseEffect = 3 then
		AddNoise()
	end if

	'write text
	for i=1 to CodeLength
		'calculate cursor pos
		CursorPos = 0
		for j=(i-1) to 1 step -1
			CursorPos = CursorPos + arrTextWidth(j) + CharTracking
		next
		CursorPos = LeftMargin + CursorPos

		'write active char
		WriteCanvas Mid(secureCode,i,1),arrTopMargin(i)
	next
end sub

sub DrawLine(x0, y0, x1, y1, valClr)
	'Reference from Donald Hearn and M. Pauline Baker, Computer Graphics C Version
	dim m, b, dx, dy

	if (NoiseEffect = 4) and (Bitmap(y0,x0) = TColor) then
		clrNoise = vbNullString
	else
		clrNoise = valClr
	end if
	Bitmap(y0,x0) = clrNoise

	dx = x1 - x0
	dy = y1 - y0
	if Abs(dx) > Abs(dy) then
		m = (dy / dx)
		b = y0 - (m * x0)

		if dx < 0 then
			dx = -1
		else
			dx = 1
		end if

		do while x0 <> x1
			x0 = x0 + dx

			if (NoiseEffect = 4) and (Bitmap(Round((m * x0) + b),x0) = TColor) then
				clrNoise = vbNullString
			else
				clrNoise = valClr
			end if
			Bitmap(Round((m * x0) + b),x0) = clrNoise
		loop
	elseif dy <> 0 then
		m = (dx / dy)
		b = x0 - (m * y0)

		if dy < 0 then
			dy = -1
		else
			dy = 1
		end if

		do while y0 <> y1
			y0 = y0 + dy

			if (NoiseEffect = 4) and (Bitmap(y0,Round((m * y0) + b)) = TColor) then
				clrNoise = vbNullString
			else
				clrNoise = valClr
			end if
			Bitmap(y0,Round((m * y0) + b)) = clrNoise
		loop
	end if
end sub

sub AddNoise()
	dim median, i, j, x0, y0, x1, y1, dx, dy, dxy

	if NoiseEffect = 1 then
		clrNoise = vbNullString
	else
		clrNoise = NColor
	end if

	for i=1 to NoiseLine
		x0 = RndInterval(1,ImageWidth)
		y0 = RndInterval(1,ImageHeight)
		x1 = RndInterval(1,ImageWidth)
		y1 = RndInterval(1,ImageHeight)

		'Check minimum line length
		dx = Abs(x1 - x0)
		dy = Abs(y1 - y0)
		median = Round(Sqr((dx * dx) + (dy * dy))/2)
		if median < MinLineLength then
			dxy = MinLineLength - median

			if x1 < x0 then
				dx = -1
			else
				dx = 1
			end if

			if y1 < y0 then
				dy = -1
			else
				dy = 1
			end if

			for j=1 to dxy
				if ((x1 + dx) < 1) or ((x1 + dx) > ImageWidth) or ((y1 + dy) < 1) or ((y1 + dy) > ImageHeight) then
					exit for
				end if
				x1 = x1 + dx
				y1 = y1 + dy
			next
		end if

		'Draw noise line
		DrawLine x0,y0,x1,y1,clrNoise
	next
end sub

function FormatHex(byval valHex,byval fixByte,fixDrctn,valReverse)
	fixByte = fixByte * 2
	tmpLen = Len(valHex)
	if fixByte > tmpLen then
		tmpFixHex = String((fixByte - tmpLen),"0")
		if fixDrctn = 1 then
			valHex = valHex & tmpFixHex
		else
			valHex = tmpFixHex & valHex
		end if
	end if

	if valReverse = true then
		tmpHex = vbNullString
		for cFrmtHex=1 to Len(valHex) step 2
			tmpHex = Mid(valHex,cFrmtHex,2) & tmpHex
		next
		FormatHex = tmpHex
	else
		FormatHex = CStr(valHex)
	end if
end function

sub SendHex(valHex)
	for cHex = 1 to Len(valHex) step 2
		Response.BinaryWrite ChrB(CByte("&H" & Mid(valHex,cHex,2)))
	next
end sub

sub SendBitmap()
	if (ImageWidth mod 4) <> 0 then
		BmpEndLine = String((4-(ImageWidth mod 4))*2,"0")
	else
		BmpEndLine = vbNullString
	end if
	BmpInfoHeader = Array("28000000","00000000","00000000","0100","0800","00000000","00000000","120B0000","120B0000","00000000","00000000")
	BmpInfoHeader(1) = FormatHex(Hex(ImageWidth),4,0,true)
	BmpInfoHeader(2) = FormatHex(Hex(ImageHeight),4,0,true)
	BmpInfoHeader(6) = FormatHex(Hex((ImageHeight * ImageWidth) + (ImageHeight * (Len(BmpEndLine) / 2))),4,0,true)
	BmpInfoHeader(9) = FormatHex(Hex(Len(BmpColorMap)/8),4,0,true)
	BmpInfoHeader(10) = BmpInfoHeader(9)
	BmpHeader = Array("424D","00000000","0000","0000","00000000")
	BmpHeader(1) = FormatHex(Hex((Len(Join(BmpHeader,"")) / 2) + (Len(Join(BmpInfoHeader,"")) / 2) + (Len(BmpColorMap) / 2) + (ImageHeight * ImageWidth) + (ImageHeight * (Len(BmpEndLine) / 2))),4,0,true)
	BmpHeader(4) = FormatHex(Hex((Len(Join(BmpHeader,"")) / 2) + (Len(Join(BmpInfoHeader,"")) / 2) + (Len(BmpColorMap) / 2)),4,0,true)

	Response.Clear
	Response.Buffer = True
	Response.ContentType = "image/bmp"
	Response.AddHeader "Content-Disposition", "inline; filename=captcha.bmp"
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

	SendHex(Join(BmpHeader,""))
	SendHex(Join(BmpInfoHeader,""))
	SendHex(BmpColorMap)
	for y=ImageHeight to 1 step -1
		for x=1 to ImageWidth
			tmpHex = Bitmap(y,x)
			if tmpHex = vbNullString then
				SendHex(BColor)
			else
				SendHex(tmpHex)
			end if
		next
		SendHex(BmpEndLine)
	next
	Response.Flush
end sub
%>

<%
'#Generate captcha
secureCode = CreateGUID(CodeLength)
Session("CAPTCHA") = secureCode
PrepareBitmap(secureCode)
if (NoiseEffect > 0) and (NoiseEffect <> 3) then
	AddNoise()
end if
SendBitmap()
%>
