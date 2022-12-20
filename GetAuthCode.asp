<%
	'明確變量申明
	Option Explicit
	Session.codepage = 65001	
	Session.Timeout = 3

	Randomize	
	
	Dim jpeg
	Dim pixelsAcross,textColour,CodeTotal
	Dim sessionnaem 
	Dim CodePicWidth,CodePicHeight
	Dim randomNumber : randomNumber = Int(Rnd * 7)+1
	
	' 驗證碼變數
	sessionnaem = "KmcAuthCode"
	
	' 驗證碼個數
	CodeTotal = 6
	
	' 驗證碼圖片寬度
	CodePicWidth = 220
	
	' 驗證碼圖片高度
	CodePicHeight = 50
	

	' 文字顏色
	textColour =  randomFomtcolor(randomNumber)
	
	' 驗證碼字距
	pixelsAcross = Int(Rnd * 40)+3
	pixelsAcross = 5
	
	Set jpeg = Server.CreateObject("Persits.jpeg")
	
	' 隨機背景圖片
	drawBackGroud randomNumber,CodePicWidth,CodePicHeight

	' 繪製文字
	doString
	
	' 隨機線
	drawLines

	' 隨機圓
	drawCircle

	' 隨機矩形
	'drawBar

	' 返回二進制，申明此為圖片
	jpeg.SendBinary
	
	Set jpeg = Nothing 


	' 函數(drawBackGroud)打開背景圖
	Function drawBackGroud(srandom,swidth,sheight)
		Jpeg.Open Server.MapPath("images/background/background"&srandom&".jpg")
		Jpeg.Width = swidth
		Jpeg.Height = sheight	
	End Function
	
	' 函數(drawLines)繪製隨機線
	Sub drawLines
		jpeg.Canvas.Pen.Color = &HADCD3C
		jpeg.Canvas.DrawLine 0, Int(Rnd * jpeg.Height), jpeg.Width, Int(Rnd * jpeg.Height)
	End Sub
	
	' 函數(drawBar)繪製隨機矩形框
	Sub drawBar
		jpeg.Canvas.Brush.Solid = False '填充
		'矩形邊框顏色
		jpeg.Canvas.Pen.Color = &H9CCF00
		'繪製矩形框
		jpeg.Canvas.Bar Int(Rnd * jpeg.Width), Int(Rnd * jpeg.Height), Int(Rnd * 50)+20,Int(Rnd * 50)+20
	End Sub

	' 函數(drawCircle)繪製隨機圓
	Sub drawCircle
		jpeg.Canvas.Brush.Solid = False '填充
		jpeg.Canvas.Pen.Color = &H8080FF
		jpeg.Canvas.Circle Int(Rnd * jpeg.Width), Int(Rnd * jpeg.Height), Int(Rnd * 10)+5
		jpeg.Canvas.Pen.Color = &HEEEEEE
		jpeg.Canvas.Circle Int(Rnd * jpeg.Width), Int(Rnd * jpeg.Height), Int(Rnd * 10)+10
	End Sub

	' 函數(doString)繪製驗證碼字串
	Sub doString
		Dim theString
		Dim x
	
		' 獲取隨機字串
		theString = createRandomString()
		
		' 
		For x = 1 to len(theString)

			' 在驗證碼圖片SHOW出字串
			addLetter Mid(theString, x, 1)
			
		Next

	End Sub

	' 函數(addLetter）在驗證碼圖片SHOW出字串
	Sub addLetter(theLetter)	
	
		' 字體顏色
		jpeg.Canvas.Font.Color = textColour

		' 字體陰影
		jpeg.Canvas.Font.ShadowColor = &HFFFFFF
			
		' 是否為粗體　加粗效果更好，故不做隨機判斷，而是直接設定加粗
		'if doTextStyle then
			jpeg.Canvas.Font.Bold = True
		'End If
		
		' 是否增加下畫線 
		'if doTextStyle then
		'	jpeg.Canvas.Font.Underlined  = True
		'End If	
		
		' 是否為斜體
		if doTextStyle then
			jpeg.Canvas.Font.Italic   = True
		End If		
		
		' 字體
		jpeg.Canvas.Font.Family = "Arial Black"'randomFont()		
		
		' 字體大小
		jpeg.Canvas.Font.Size = randomFontSize()
		
		' 文字清晰度
		jpeg.Canvas.Font.Quality = 4
		
		' 背景色　因使用隨機背景圖、註解掉
		'jpeg.Canvas.Font.BkColor = backColour
		
		' 字體背景模式(處理平滑)
		jpeg.Canvas.Font.BkMode = "transparent"
		
		' 繪製字串
		jpeg.canvas.print pixelsAcross, Int(Rnd * 5), theLetter
		
		' 文字寬度
		pixelsAcross = pixelsAcross + Int(Rnd * 10)+30
		
	End Sub
	
	' 返回隨機值50%
	Function doTextStyle()
		if Rnd() > 0.5 then
			doTextStyle = true
		else
			doTextStyle = false
		end if
	End Function

	' 返回驗證碼中各文字的大小
	Function randomFontSize()
		Dim theNumber
		' 獲取一個隨機大小，範圍(40-60)
		theNumber = Int(Rnd * 20) + 40
		randomFontSize = theNumber
		
	End Function

	' 返回隨機驗證碼文字顏色
	Function randomFomtcolor(srandomm)
		Dim arrFomtcolor(8)
		arrFomtcolor(1) = &HBDE3FF
		arrFomtcolor(2) = &HD68618
		arrFomtcolor(3) = &H086529
		arrFomtcolor(4) = &H637594
		arrFomtcolor(5) = &Hffffff
		arrFomtcolor(6) = &HBDDBF7
		arrFomtcolor(7) = &H08756B
		arrFomtcolor(8) = &H295131
		randomFomtcolor = arrFomtcolor(srandomm)
	End Function 
	
	' 返回隨機字體
	Function randomFont()
		Dim theNumber	
		Dim font	
		' 取得1-6區間內一隨機字元
		theNumber = Int(Rnd * 5) + 1
		' 隨機字體
		if theNumber =1 then
			font = "Arial Black"
		elseif theNumber =2 then
			font = "Courier New"
		elseif theNumber =3 then
			font = "Helvetica"
		elseif theNumber =4 then
			font = "Times New Roman"
		elseif theNumber =5 then
			font = "Verdana"
		else
			font = "Geneva"
		end If
		randomFont = font
	
	End Function

	' 返回隨機驗證碼字元
	Function createRandomString
		Dim outputString
		Dim x
        For x = 0 To CodeTotal-1
			' 英文字元出現機率60%, 數字出現機率40%
			'if rnd() < 0.6 then
				' 返回一隨機英文字元
            			'outputString = outputString & Chr(Int((26 * rnd()) + 65))
			'else
				' 返回一隨機數字字元
				outputString = outputString & Chr(Int((10 * rnd()) + 48))
			'end if
        Next
		Session(sessionnaem) = outputString
        createRandomString = outputString	
	End Function


	
%>

