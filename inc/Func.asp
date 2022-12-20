<%
Function FunSQL(funsqlstr)
	IF funsqlstr&"">"" Then
		FunSQL = FunCheckText(trim(funsqlstr))
	ELSE
		FunSQL = NULL
	END IF
END Function

Function FunCheckText(fnsqlvalue)
	fnstr = replace(fnsqlvalue,"""","”")
	'fnstr = replace(fnstr,"'","''")
	FunCheckText = fnstr
end Function


'檢查是否有使用', ", %等字元
Function checkhack(keystr)
	checkstr = replace(keystr,"'","")
	checkstr = replace(checkstr,"%","")
	checkhack = checkstr
End Function


Function checkaccount(keystr)
	checkstr = replace(keystr,"~","")
	checkstr = replace(checkstr,"!","")
	checkstr = replace(checkstr,"@","")
	checkstr = replace(checkstr,"#","")
	checkstr = replace(checkstr,"$","")
	checkstr = replace(checkstr,"%","")
	checkstr = replace(checkstr,"^","")
	checkstr = replace(checkstr,"&","")
	checkstr = replace(checkstr,"*","")
	checkstr = replace(checkstr,"(","")
	checkstr = replace(checkstr,")","")
	checkstr = replace(checkstr,"_","")
	checkstr = replace(checkstr,"+","")
	checkstr = replace(checkstr,"|","")
	checkstr = replace(checkstr,"-","")
	checkstr = replace(checkstr,"=","")
	checkstr = replace(checkstr,"[","")
	checkstr = replace(checkstr,"]","")
	checkstr = replace(checkstr,"{","")
	checkstr = replace(checkstr,"}","")
	checkstr = replace(checkstr,"'","")
	checkstr = replace(checkstr,"""","")
	checkstr = replace(checkstr,":","")
	checkaccount = checkstr
End Function



Function XSSFilter(STR)
	TempXSSSTR = STR
	DebugSTR=	"'" & "@" & _
				"<" & "@" & _
				">" & "@" & _
				"&" & "@" & _
				"%" & "@" & _
				"alert" & "@" & _
				"script"
			   
	DebugArray = split(DebugSTR,"@")
	For i=0 to Ubound(DebugArray)
		TempXSSSTR = Replace(UCASE(TempXSSSTR),UCASE(DebugArray(i)),"")
	Next
	XSSFilter = TempXSSSTR
End Function



Function GetNewID(TableName,fieldName)
	fieldstr = "Max("&fieldName&")+1"
	MaxID = GetFieldValue(TableName,"",fieldstr)
	IF ISNULL(MaxID) OR MaxID="" Then
		GetNewID=1
	ELSE
		GetNewID=cint(MaxID)
	End IF
End Function



Function GetFieldValue(TableName,WhereStr,FieldName)
	'ConnNM:使用之UDL名稱
	'TableName:資料表格名稱
	'wherestr:條件式
	'FieldName:欲取得欄位
	'Example:SerialID=GetFieldValue("web_eip","afs_flow","SerialID='abcde'","Status")
	Set FunOBJconn = Server.CreateObject("ADODB.Connection")
	FunOBJconn.Open DBConnSTR
	Set FunRS = Server.CreateObject("ADODB.RecordSet")
	IF WhereStr="" Then
		FinalWhereStr = ""
	ELSE
		FinalWhereStr = " WHERE " & WhereStr
	End IF
 	FunStr = "SELECT " & FieldName & " FROM " & TableName & FinalWhereStr
	'response.write funstr
	set FunRS = FunOBJconn.execute(FunStr)
	if not FunRS.eof then
		GetFieldValue = FunRS(0)
	else
		GetFieldValue = ""
	end if

	FunRS.Close
	FunOBJconn.Close
	Set FunRS = nothing
	Set FunOBJconn = nothing
End Function

Function ExcuteSQL(SQLstring)
	Set FunOBJconn = Server.CreateObject("ADODB.Connection")
	FunOBJconn.Open DBConnSTR
	FunOBJconn.Execute(SQLstring)
	FunOBJconn.Close
	Set FunOBJconn = nothing
End Function





Function uSendHttpMail(ToEmlAddress,FromEmlAddress,MailTitle,FunUrls) 
	Set iMsg = Server.CreateObject("CDO.Message") 
	'一定要加上這一段，否則如果你的Web Server上有安裝 Outlook 2002 等版本更新了CDOEX.DLL 
	'在ASP中會導致 senduse 的錯誤 
	Set iConf = CreateObject("CDO.Configuration") 
	Set Flds = iConf.Fields
	'取得SMTP SERVER
	SMTPSERVER = GetFieldValue("STATUS_DESC","STATUS=2 AND STATUS_TYPE=11","NOTES")
	'取得帳號
	SMTPID = GetFieldValue("STATUS_DESC","STATUS=3 AND STATUS_TYPE=11","NOTES")	
	'取得密碼 
	SMTPPW = GetFieldValue("STATUS_DESC","STATUS=4 AND STATUS_TYPE=11","NOTES")
	'取得通訊埠 
	SMTPPORT = GetFieldValue("STATUS_DESC","STATUS=5 AND STATUS_TYPE=11","NOTES")
	'取得SSL認證
	SMTPSSL =  GetFieldValue("STATUS_DESC","STATUS=7 AND STATUS_TYPE=11","NOTES")
	
	
	With Flds 
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1 
		'.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = "c:\inetpub\mailroot\pickup"
		IF trim(SMTPID&"")>"" Then
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTPID
		End IF
		
		IF trim(SMTPPW&"")>"" Then
		.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTPPW
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = "1"
		End IF
		'.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "12142"
		'.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "242424"

		.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") =2 
		'遠端SMTP主機名稱或IP位址 
		.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver" )= SMTPSERVER
		
		'遠端SSL認證
		IF SMTPSSL&"">"" and SMTPSSL="true" Then
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
		End IF
		
		'遠端SMTP主機埠號 Server port 
		.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPPORT 
		.Update 
	End With 

	MailID = Rdn12() & Rdn12()
	call WriteMailSendHistory(MailID,ToEmlAddress,FromEmlAddress,MailTitle,"",FunUrls)
	
	With iMsg
		.Configuration=iConf 
		'一定要指明 Encoding 為 7bits，否則送HTML格式的 email 時，因為編碼的原因，會導致 .jpg .不見 
		.BodyPart.ContentTransferEncoding = "7bit" 
		'如果使用多國語言時，才要設定 Charset 
		'Mail1.BodyPart.Charset = "big5" 
		.To = ToEmlAddress 
		.From = FromEmlAddress 
		.Subject = MailTitle 
		'.HtmlBody = MailContext 
		'發送網站上的內容 or 本機端的htm檔案 
		'example: 
		'SiteUrl = GetFieldValue("STATUS_DESC","STATUS=1 AND STATUS_TYPE=11","NOTES")
		'ServerUrl = GetServerUrl()
		'FinalFunUrls = replace(FunUrls,SiteUrl,ServerUrl)
		FinalFunUrls = FunUrls
		'response.write FinalFunUrls
		'response.end
		.CreateMHTMLBody FinalFunUrls '遠端網站
		'.CreateMHTMLBody "file://c:/picts/test.htm" '本機檔案
		'.CreateMHTMLBody MailContext
		.Send
	End With 
	
	'成功回寫
	call DoneMailSendHistory(MailID)	
	
	Set iMsg=Nothing 
	Set iConf = Nothing 
End Function 


'取得 zh.txt中的說明欄位
Function GetText(findstr)
    Dim objStream
    Set objStream = Server.CreateObject("ADODB.Stream")
    With objStream
        .Type = 2
        .Mode = 3
        .Open
        .LoadFromFile Server.MapPath("inc/zh.txt")
        If Err.Number<>0 Then
			Response.Write "Error"
			Err.Clear
			Response.End
        End If
        .Charset = "big5"
        .Position = 2
		TextBody = .ReadText
		.Close
    End With
    Set objStream = Nothing
	IF instr(TextBody,findstr)>0 Then
		text1=right(Textbody,len(TextBody)-instr(TextBody,findstr)+1)
		text1=left(text1,instr(text1,vbcrlf)-1)
		text1=right(text1,len(text1)-instr(text1,"="))
		GetText=replace(text1,"\n","<br>")
	ELSE
		GetText=""
	End IF
End Function 
	

'取得IE的網址
Function GetIEurl() 
  Dim strTemp 
If LCase(Request.ServerVariables("HTTPS")) = "off" Then 
strTemp = "http://" 
Else 
strTemp = "https://" 
  End If 
  strTemp = strTemp & Request.ServerVariables("SERVER_NAME") 
  If Request.ServerVariables("SERVER_PORT") <> 80 Then strTemp = strTemp & ":" & Request.ServerVariables("SERVER_PORT") 
  strTemp = strTemp & Request.ServerVariables("URL") 
  If Trim(Request.QueryString) <> "" Then strTemp = strTemp & "?" & Trim(Request.QueryString) 
  GetIEurl = strTemp 
End Function

Function EarseIEurl(urlstr,param)
	IF instr(urlstr,"?")>0 Then
		iepars = right(urlstr,len(urlstr)-instr(urlstr,"?"))
		If instr(iepars,param&"=")>0 Then
			pars1=left(iepars,instr(iepars,param&"=")-2)
			pars2=right(iepars,len(iepars)-(len(pars1)+1))
			'response.write pars2&"<br>"
			IF instr(pars2,"&")>0 then
				pars2=right(pars2,len(pars2)-instr(pars2,"&"))
				pars3 = pars1 &"&"& right(pars2,len(pars2)-instr(pars2,"&"))
			ELSE
				pars3 = pars1
			End IF
			EarseIEurl = "?"&pars3
		ELSE
			EarseIEurl = "?"&iepars
		End IF
	ELSE
		EarseIEurl = "?auex="
	END IF
End Function


Function MakeComboCon(connname,SQLstr,textField,valueField,selectvalue)
	Set FunOBJconn = Server.CreateObject("ADODB.Connection")
	SELECT CASE connname
		CASE "SA"
			UseDBConnSTR = DBConnSTR	
		CASE "LISA"
			UseDBConnSTR = LISADBConnSTR
	END SELECT
	FunOBJconn.Open UseDBConnSTR
	Set FunRS = Server.CreateObject("ADODB.RecordSet")
	FunRS.Open SQLstr,FunOBJconn,3,1
		IF FunRS.recordcount > 0 Then
			FunRS.MoveFirst
			Do while not FunRS.eof
				response.write "<Option value="""&FunRS(valueField)&""""
				IF cstr(selectvalue) = cstr(FunRS(valueField)) Then
					response.write " selected "
				End IF
				response.write ">"&FunRS(textField)&"</Option>" & vbcrlf
				FunRS.MoveNext
			Loop
		End IF
	FunRS.Close
	FunOBJconn.Close
	Set FunRS = nothing
	Set FunOBJconn = nothing
End Function



Function MakeCombo(SQLstr,textField,valueField,selectvalue)
	IF selectvalue&""="" Then
		selectvalue = ""
	end IF
	Set FunOBJconn = Server.CreateObject("ADODB.Connection")
	FunOBJconn.Open DBConnSTR
	Set FunRS = Server.CreateObject("ADODB.RecordSet")
	FunRS.Open SQLstr,FunOBJconn,3,1
		IF FunRS.recordcount > 0 Then
			FunRS.MoveFirst
			Do while not FunRS.eof
				FunValue = FunRS(valueField)&""
				response.write "<Option value="""&FunRS(valueField)&""""
				IF selectvalue&"">"" Then
					IF cstr(selectvalue) = cstr(FunValue) Then
						response.write " selected "
					End IF
				End IF
				response.write ">"&FunRS(textField)&"</Option>" & vbcrlf
				FunRS.MoveNext
			Loop
		End IF
	FunRS.Close
	FunOBJconn.Close
	Set FunRS = nothing
	Set FunOBJconn = nothing
End Function



Function MakeComboStr(Combostr,unit)
	Str = split(Combostr,",")
	For i = 0 to Ubound(Str)
		IF i=Ubound(str) Then
			addstr = "以上"
		end IF
		response.write "<Option value="""&Str(i)&""">"&Str(i)&" "&unit&""&addstr&"</Opton>"
	Next
End Function


Function MakeTextCombo(Combostr,defaultstr)
	IF defaultstr&""="" Then
		defaultstr = ""
	ELSE
		defaultstr = cstr(defaultstr)
	End IF
	Str = split(Combostr,",")
	For i = 0 to Ubound(Str)
		response.write "<Option value="""&Str(i)&""""
		IF cstr(Str(i))=defaultstr then
		response.write " selected "
		End IF
		response.write ">"&Str(i)&"</Opton>" & vbcrlf
	Next
End Function


Function MakeTextValueCombo(ComboTextStr,ComboValueStr,defaultstr)
	FunTextSTr = split(ComboTextStr,",")
	FunValueStr = split(ComboValueStr,",")
	IF Ubound(FunTextSTr)<>Ubound(FunValueStr) Then
		response.write "Text 與 Value 數量不相符。"
		response.end
	End IF
	For i = 0 to Ubound(FunTextSTr)
		response.write "<Option value="""&FunValueStr(i)&""""
		IF FunValueStr(i)=defaultstr then
		response.write " selected "
		End IF
		response.write ">"&FunTextSTr(i)&"</Opton>" & vbcrlf
	Next
End Function

'產生12位亂數數字
Function Rdn12()

	'使亂數產生時不照順序
	Randomize
	
	'亂數產生十二碼位數
	'方法：分六次產生，每次產生一個不同的亂數，並加上當時秒數。
	Code = Right("00" & Second(Now) + Int((39 * Rnd)+1),2)
	Code = Code & Right("00" & Second(Now) + Int((39 * Rnd)+1) ,2)
	Code = Code & Right("00" & Second(Now) + Int((39 * Rnd)+1) ,2)
	Code = Code & Right("00" & Second(Now) + Int((39 * Rnd)+1) ,2)
	Code = Code & Right("00" & Second(Now) + Int((39 * Rnd)+1) ,2)
	Code = Code & Right("00" & Second(Now) + Int((39 * Rnd)+1) ,2)
	
	Rdn12 = Code
	
	'將十二位數拆成十二組，兩位數一組。
	For I = 1 To 12
	Session("Temp_" & I) = Clng(Asc(Mid(Code,I,1)))
	Next
	
	'設定十二個變數，內容為上面拆出來的變數照順序合併後，再乘上一組一到一千的亂數，再取三十六的於數。
	Session("Code_1") = ((Session("Temp_1") & Session("Temp_2")) * Int((1000 * Rnd)+1)) Mod 36
	Session("Code_2") = ((Session("Temp_2") & Session("Temp_3")) * Int((1000 * Rnd)+1)) Mod 36
	Session("Code_3") = ((Session("Temp_3") & Session("Temp_4")) * Int((1000 * Rnd)+1)) Mod 36
	Session("Code_4") = ((Session("Temp_4") & Session("Temp_5")) * Int((1000 * Rnd)+1)) Mod 36
	Session("Code_5") = ((Session("Temp_5") & Session("Temp_6")) * Int((1000 * Rnd)+1)) Mod 36
	Session("Code_6") = ((Session("Temp_6") & Session("Temp_7")) * Int((1000 * Rnd)+1)) Mod 36
	Session("Code_7") = ((Session("Temp_7") & Session("Temp_8")) * Int((1000 * Rnd)+1)) Mod 36
	Session("Code_8") = ((Session("Temp_8") & Session("Temp_9")) * Int((1000 * Rnd)+1)) Mod 36
	Session("Code_9") = ((Session("Temp_9") & Session("Temp_10")) * Int((1000 * Rnd)+1)) Mod 36
	Session("Code_10") = ((Session("Temp_10") & Session("Temp_11")) * Int((1000 * Rnd)+1)) Mod 36
	Session("Code_11") = ((Session("Temp_11") & Session("Temp_12")) * Int((1000 * Rnd)+1)) Mod 36
	Session("Code_12") = ((Session("Temp_12") & Session("Temp_1")) * Int((1000 * Rnd)+1)) Mod 36
	
	'將上列十二組變數合併，並用 Chr 函式轉為 0-9 A-Z 英數混合的十二碼字串。
	For I = 1 To 12
		If Session("Code_" & I) > 9 Then
			Session("Code_" & I) = Chr(Session("Code_" & I) + 55)
		End If
		Password = Password & Session("Code_" & I)
	Next
	
	Rdn12 = Password
	
End Function



function TNum(number)
	if number&"">"" Then
		numberaf = CDbl(number)
	end if
	if not(isnumeric(numberaf)) or cdbl(numberaf) = 0 then
		result = 0
		elseif len(fix(numberaf)) < 4 then
		result = number
	else
		dim pos,res,loopcount,tempresult,dec,result,funflag
		if numberaf<0 then
			numberaf = numberaf*-1
			funflag=true
		end if
		pos = instr(1,numberaf,".")
		if pos > 0 then
			dec = mid(numberaf,pos)
		end if
		res = strreverse(fix(numberaf))
		loopcount = 1
		while loopcount <= len(res)
			tempresult = tempresult + mid(res,loopcount,3)
			loopcount = loopcount + 3
			if loopcount <= len(res) then
			tempresult = tempresult + ","
			end if
		wend
		result = strreverse(tempresult) + dec
	end if
	if funflag then
		numberaf = numberaf*-1
		result = "-" & result
	end if
	TNum = result
end function






'十六進位字串轉換成十進位數字
Function hex2dec(str)
	totals = 0
	newstr = UCASE(str)
	For i=0 to 5
		nowstr = right(newstr,1)
		SELECT CASE nowstr
			CASE "0"
				num = 0
			CASE "1"
				num = 1
			CASE "2"
				num = 2
			CASE "3"
				num = 3
			CASE "4"
				num = 4
			CASE "5"
				num = 5
			CASE "6"
				num = 6
			CASE "7"
				num = 7
			CASE "8"
				num = 8
			CASE "9"
				num = 9
			CASE "A"
				num = 10
			CASE "B"
				num = 11
			CASE "C"
				num = 12
			CASE "D"
				num = 13
			CASE "E"
				num = 14
			CASE "F"
				num = 15
		END SELECT
		IF i>0 then
			for x=1 to i
				base = base * 16
			Next
		else
			base = 1
		End IF
		totals = totals + num*base
		IF i<5 Then
			newstr = left(newstr,5-i)
		End IF
		base = 1
	Next
	hex2dec = totals
End Function


Function AddOther(otherstr,addstr,tabstr)
	IF otherstr="" Then
		AddOther = addstr
	ELSE
		AddOther = otherstr & tabstr & addstr
	End IF
End Function


Function thisDate()
	FunYY = Year(now())
	FunMM = Month(now())
	FunDD = Day(now())
	thisDate = CDATE(FunYY&"/"&FunMM&"/"&FunDD)
End Function

Function thisDateTime()
	thisDateTime = DateFormat(now(),6)
End Function

Function DateFormat(dt,dtstr)
	IF dt&"">"" Then
		IF NOT ISNULL(dt) AND ISDate(dt) Then
			dt = CDATE(dt)
			FunYY = Year(dt)
			FunMM = Month(dt)
			FunDD = Day(dt)
			FunHH = Hour(dt)
			FunMI = Minute(dt)
			FunSS = Second(dt)
			SELECT CASE dtstr
				CASE 1 '完整日期(含0)
					DateFormat = FunYY&"/"&Right("0"&FunMM,2)&"/"&Right("0"&FunDD,2)
				CASE 2 '簡短日期(省略0)
					DateFormat = FunYY&"/"&FunMM&"/"&FunDD
				CASE 3 '民國年(含0)
					DateFormat = FunYY-1911&"/"&Right("0"&FunMM,2)&"/"&Right("0"&FunDD,2)
				CASE 4 '民國年(省略0)
					DateFormat = FunYY-1911&"/"&FunMM&"/"&FunDD
				CASE 5 'YY/MM/DD HH:MI
					DateFormat = FunYY-1911&"/"&Right("0"&FunMM,2)&"/"&Right("0"&FunDD,2) & " " & Right("0"&FunHH,2) & ":" & Right("0"&FunMI,2)
				CASE 6 'YYYY/MM/DD HH:MI:SS
					DateFormat = FunYY&"/"&Right("0"&FunMM,2)&"/"&Right("0"&FunDD,2) & " " & Right("0"&FunHH,2) & ":" & Right("0"&FunMI,2)&":"&Right("0"&FunSS,2)
				CASE 7 'YYYY/MM/DD HH:MI
					DateFormat = FunYY&"/"&Right("0"&FunMM,2)&"/"&Right("0"&FunDD,2) & " " & Right("0"&FunHH,2) & ":" & Right("0"&FunMI,2)
				CASE 8 'HH:MI
					DateFormat = Right("0"&FunHH,2) & ":" & Right("0"&FunMI,2)
				CASE 9 'MM/DD
					DateFormat = Right("0"&FunMM,2)&"/"&Right("0"&FunDD,2)
				CASE 10 'HHMI
					DateFormat = Right("0"&FunHH,2) & Right("0"&FunMI,2)	
				CASE 11 '民國年無/
					DateFormat = Right("000"&FunYY-1911,3)&Right("0"&FunMM,2)&Right("0"&FunDD,2)
			End SELECT
		ELSE
			DateFormat = ""
		End IF
	ELSE
		DateFormat = ""
	END IF
End Function

Function Dt2Int(dtstr)
	Dt2Int = replace(DateFormat(dtstr,1),"/","")
End Function

Function Int2Dt(intstr)
	IF intstr&"">"" Then
		Int2Dt = CDATE(left(intstr,4)&"/"&Right(left(intstr,6),2)&"/"&Right(intstr,2))
	ELSE
		Int2Dt = ""
	END IF
End Function

Function BackIcon()
	BackIcon="<img src=""Images/back.gif"" border=""0"" onclick=""window.history.back();return false;"" style=""cursor:hand;"" alt=""回上頁"">"
End Function

Function dbstr(str,datatype)
	IF str&"" = "" Then
		dbstr = " NULL "
	ELSE	
		SELECT CASE datatype
			CASE "text"
				dbstr="'"&str&"'"
			CASE "num"
				dbstr=str
			CASE "date"
				dbstr="'"&str&"'"
		END SELECT
	End IF
End Function

Function ErrorBack(errstr)
	response.write "<Script>alert("""&errstr&""");"
	response.write "history.back();"
	response.write "</Script>"
	response.end
END Function

Function ErrorClose(errstr)
	response.write "<Script>alert("""&errstr&""");"
	response.write "window.close();;"
	response.write "</Script>"
	response.end
END Function

Function ErrorAlert(errstr)
	response.write "<Script>alert("""&errstr&""");"
	response.write "</Script>"
	response.end
END Function


Function GetMaxSerial(tablestr,wherestr,fieldname)
	Set FunOBJconn = Server.CreateObject("ADODB.Connection")
	FunOBJconn.Open DBConnSTR
	Set FunRS = Server.CreateObject("ADODB.RecordSet")
	IF wherestr<>"" Then
		FunWhereSTR = " WHERE " & wherestr
	ENd IF
	FunSQLSTR = "SELECT MAX("&fieldname&") FROM "&tablestr & FunWhereSTR
	FunRS.Open FunSQLSTR,FunOBJconn,3,1
		IF FunRS.recordcount > 0 Then
			IF FunRS(0)&"">"" Then
				GetMaxSerial = cint(FunRS(0))+1
			ELSE
				GetMaxSerial = 1
			END IF
		ELSE
			GetMaxSerial = 1
		END IF
	FunRS.Close
	FunOBJconn.Close
	Set FunRS = nothing
	Set FunOBJconn = nothing
END Function

Function GetPrintNo(NOSTR,FunN)
	Set FunOBJconn = Server.CreateObject("ADODB.Connection")
	FunOBJconn.Open DBConnSTR
	Set FunRS = Server.CreateObject("ADODB.RecordSet")
	Set FunRS2 = Server.CreateObject("ADODB.RecordSet")
	IF UCASE(Left(NOSTR,2))="SV" Then
		FunField = "PRE_NO"
	ELSE
		FunField = "PRE_NO"
	END IF
	
	FunSQLSTR = "SELECT CNT+1,UNIQID FROM CASE_Bpart_CNT WHERE " & FunField &"='"&NOSTR&"' AND STATUS=" & FunN
	FunRS.Open FunSQLSTR,FunOBJconn,3,1
	application.lock
		IF FunRS.recordcount > 0 Then
			FunCNT = cint(FunRS(0))
			FunUNIQID = FunRS(1)
			
			FunRS2.open "SELECT * FROM CASE_Bpart_CNT WHERE UNIQID="&FunUNIQID,FunOBJconn,2,3
			FunRS2("UPDATE_DT")=thisDateTime()
			FunRS2("UPDATE_USER")=Session("UID")
		ELSE
			FunCNT = 1
			
			FunRS2.open "CASE_Bpart_CNT", FunOBJconn, 2, 3
			FunRS2.addnew
			FunRS2(FunField)=UCASE(NOSTR)
			FunRS2("STATUS")=FunN
			FunRS2("CREATE_DT")=thisDateTime()
			FunRS2("CREATE_USER")=Session("UID")			
		END IF		
		FunRS2("CNT")=FunCNT
		FunRS2.update
		FunRS2.close
	application.unlock

	FunRS.Close
	FunOBJconn.Close
	Set FunRS = nothing
	Set FunRS2 = nothing
	Set FunOBJconn = nothing
	GetPrintNo = FunCNT
END Function


Function CaseViewLog(FunCaseNo)
	Set FunOBJconn = Server.CreateObject("ADODB.Connection")
	FunOBJconn.Open DBConnSTR
	Set FunRS = Server.CreateObject("ADODB.RecordSet")
	application.lock
		FunRS.open "Case_ViewLog", FunOBJconn, 2, 3
		FunRS.addnew
		FunRS("CASE_NO")=FunCaseNo
		FunRS("VIEW_DT")=thisDateTime()
		FunRS("VIEW_USER")=Session("UID")
		FunRS.update
		FunRS.close
	application.unlock
	FunOBJconn.Close
	Set FunRS = nothing
	Set FunOBJconn = nothing
End Function


 function showAllrequest()  '可淂知所有Request值
    tmpStr = "<TABLE border=""1"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#111111""><TR><TH colspan=2>Request.Form</TH></TR><TR><TH bgcolor=""#AAAA""><B>Key</B></TD><TH bgcolor=""#AAAA""><B>Value</B></TD></TR>" & vbCrLf
    for each Key in request.Form
      tmpStr = tmpStr & "<TR><TD>" & Key & "</TD><TD>" & request.Form(Key) & "</TD></TR>" & vbCrLf
    next

    tmpStr = tmpStr & "</TABLE>" & vbCrLf
    response.write tmpStr&"<BR>"
    tmpStr = "<TABLE border=""1"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#111111""><TR><TH colspan=2>Request.QueryString</TH></TR><TR><TH bgcolor=""#AAAA""><B>Key</B></TD><TH bgcolor=""#AAAA""><B>Value</B></TD></TR>" & vbCrLf
    for each Key in request.QueryString
      tmpStr = tmpStr & "<TR><TD>" & Key & "</TD><TD>" & request.QueryString(Key) & "</TD></TR>" & vbCrLf
    next

    tmpStr = tmpStr & "</TABLE>" & vbCrLf
    response.write tmpStr
  end function


'計算兩點距離 使用WGS84經緯度
'Exp:
' A (24.999372896504493,121.50824745298521) -> B(24.995454487110933,121.51356090169209)
' Distance = GetDistance(24.999372896504493,121.50824745298521,24.995454487110933,121.51356090169209)

Function rad(d)
	Dim Pi
	Pi = 3.1415926535898 '圓周率
	rad = d * Pi / 180
End Function


Function GetDistance(lat1,lng1,lat2,lng2) '開始緯度,開始經度
	EARTH_RADIUS = 6378.137 '地球半徑單位km
	Dim radlat1, radlat2
	Dim a , b , s , Temp 
	radlat1 = rad(lat1)
	radlat2 = rad(lat2)
	a = radlat1 - radlat2
	b = rad(lng1) - rad(lng2)
	Temp = Sqr(Sin(a / 2) ^ 2 + Cos(radlat1) * Cos(radlat2) * Sin(b / 2) ^ 2)
	s = 2 * Atn(Temp / Sqr(-Temp * Temp + 1))     'asin
	s = s * EARTH_RADIUS
	GetDistance = s
End Function


'使用中文地址取得WGS84座標 //限台灣地址使用
'回傳經緯度格式：lat@lng 
function getEarthXY(FunAddress)	
	Dim address,apikey,urls,requestText
	Dim ResultArray
	IF FunAddress&"">"" Then
		address = Server.URLEncode(FunAddress)
		apikey = "ABQIAAAAu3YNxTq2pZ5-No0-mG2LHxSshJVcHuZvwzRCZAKk5Z4egPI5vhQcrRG9xfdxUXsAnTeSxoxrNvzS6Q"
		urls = "http://maps.google.com.tw/maps/geo?q="&address&"&output=csv&key=" & apikey
	
		Set xmlhttp = CreateObject("Microsoft.XMLHTTP") 
		With xmlhttp 
			.Open "get", urls, False, "", "" 
			.Send 
			if .status<>200 then
				   getEarthXY = "0@0"
				   Set xmlhttp = Nothing 
				   Exit Function
			end if	  
			requestText = .ResponseText
		End With 
		ResultArray = split(requestText,",")
		'getEarthXY = ResultArray(0)
		IF ResultArray(0)&"">"" Then
			IF ResultArray(0)="200" Then
				getEarthXY = ResultArray(2) & "@" & ResultArray(3)
			ELSE
				getEarthXY = "0@0"
			END IF
		ELSE
			getEarthXY = "0@0"
		END IF
		Set xmlhttp = Nothing 
	ELSE
		getEarthXY = "0@0"
	END IF
End Function



Function SendMSG(FunFrom,FunTarget,FunMsgType,FunMsgTitle,FunMsgDescr)
	Set FunOBJconn = Server.CreateObject("ADODB.Connection")
	FunOBJconn.Open DBConnSTR
	Set FunRS = Server.CreateObject("ADODB.RecordSet")
	application.lock
		FunRS.open "MSG_BOX", FunOBJconn, 2, 3
		FunRS.addnew
		FunRS("Person_ID")=FunTarget
		FunRS("MSG_TYPE")=FunMsgType
		FunRS("MSG_TITLE")=FunMsgTitle
		FunRS("MSG_DESCR")=FunMsgDescr
		FunRS("read_flag")=0
		FunRS("CREATE_DT")=thisDateTime()
		FunRS("CREATE_USER")=FunFrom
		FunRS.update
		FunRS.close
	application.unlock
	FunOBJconn.Close
	Set FunRS = nothing
	Set FunOBJconn = nothing
End Function


Function CombSearchKey()
	funsearchkey = ""
	for each Key in request.Form
		funsearchkey = AddOther(funsearchkey,Key&"@@"&request.Form(Key),"#and#")
	Next
	for each Key in request.QueryString
		funsearchkey = AddOther(funsearchkey,Key&"@@"&request.QueryString(Key),"#and#")
	Next
	CombSearchKey = funsearchkey
End Function 

Function DoSearchKey(searchKey)
	ListArray = split(searchKey,"#and#")
	%>
	<Form action="index.asp" method="get" name="myform">
	<%
	For i=0 to ubound(ListArray)
		FieldArray = split(ListArray(i),"@@")
		IF FieldArray(0)<>"submit" Then
	%>
		<Input type="hidden" name="<%=FieldArray(0)%>" value="<%=FieldArray(1)%>">
	<%
		end IF
	Next
	%>
	</Form>
	<Script>
    document.myform.submit();
    </Script>
	<%
ENd Function


Function DoSearchKeyForm(searchKey)
	ListArray = split(searchKey,"#and#")
	%>
	<Form action="index.asp" method="get" name="myform">
	<%
	For i=0 to ubound(ListArray)
		FieldArray = split(ListArray(i),"@@")
		IF FieldArray(0)<>"submit" Then
	%>
		<Input type="hidden" name="<%=FieldArray(0)%>" value="<%=FieldArray(1)%>">
	<%
		end IF
	Next
	%>
	</Form>
	<%
ENd Function



Function PageForm()
	%>
    <Script>
	function switchpage(pages){
		if(pages=="") {
			document.PageForm.page.value=1;
		}else{
			document.PageForm.page.value=pages;
		}
		document.PageForm.submit();
	}
	</Script>
	<Form action="index.asp" method="get" name="PageForm">
	<%
	pageflag = true
	for each Key in request.Form
		IF UCASE(key)="PAGE" Then
			pageflag = false
		END IF
		IF UCASE(key)<>"SUBMIT" Then
		%>
		<Input type="hidden" value="<%=request.Form(Key)%>" name="<%=Key%>" >
		<%
		END IF
	Next
	for each Key in request.QueryString
		IF UCASE(key)="PAGE" Then
			pageflag = false
		END IF
		IF UCASE(key)<>"SUBMIT" Then		
		%>
		<Input type="hidden" value="<%=request.QueryString(Key)%>" name="<%=Key%>" >
		<%
		END IF
	Next
	IF pageflag Then
		%>
		<Input type="hidden" value="" name="page" >
        <%
	End IF
	%>
	</Form>
	<%
End Function 



Function downloadfile(url,fname)

	Response.Buffer=true
	
	if fname<>"" then
	  trueurl = Server.MapPath(url)
	  'trueurl=WebFolder & url
	end if
	'response.write trueurl
	'response.end
	set objFso=server.CreateObject("scripting.filesystemobject")
	set fn=objFso.GetFile(trueurl)
	flsize=fn.size
	flname=fn.name
	set fn=nothing
	set objFso=nothing
	
	set objStream=server.CreateObject("adodb.stream")
	objStream.Open 
	objStream.Type=1
	objStream.LoadFromFile trueurl
	
	select case lcase(right(flname,4))
	  case ".asf"
	  ContentType="video/x-ms-asf"
	  case ".avi"
	  ContentType="video/avi"
	  case ".doc"
	  ContentType="application/msword"
	  case ".zip"
	  ContentType="application/zip"
	  case ".xls"
	  ContentType="application/vnd.ms-excel"
	  case ".gif"
	  ContentType="image/gif"
	  case ".jpg","jpeg"
	  ContentType="image/jpeg"
	  case ".wav"
	  ContentType="audio/wav"
	  case ".mp3"
	  ContentType="audio/mpeg3"
	  case ".mpg", "mpeg"
	  ContentType="video/mpeg"
	  case ".rtf"
	  ContentType="application/rtf"
	  case ".htm","html"
	  ContentType="text/html"
	  case ".txt"
	  ContentType="text/plain"
	Case ".ASP", ".ASA", "ASPX", "ASAX", ".MDB"
		Response.Write "受保護檔案,不能下載."
		Response.End
	  case else
	  ContentType="appliation/octet-stream"
	end select
	
	Response.AddHeader "Content-Disposition", "attachment; filename="&fname
	Response.AddHeader "Content-Length", flsize
	Response.Charset="utf-8"
	Response.ContentType=ContentType
	Response.BinaryWrite objStream.Read 
	Response.Flush 
	Response.Clear()
	objStream.Close
	set objStream=nothing
   
end function


Function FunMoveFile(FunFileSource,FunFileTarget)
	
   '------------------------------------------------------------------------------
   ' 判斷檔案是否已存在 by Mark
   ' 並且移動檔案至特定目錄
   ' sample : call FunMoveFile("temp/abc.txt","objtect/txt/abc.txt")
   '------------------------------------------------------------------------------
	
	'判斷移動目錄是否存在，不存在則建立之
	IF left(FunFileTarget,1)="/" Then
		FunTargetStr = "/"
		FunFileTarget = right(FunFileTarget,len(FunFileTarget)-1)
	End IF
	
	Set FunFso = Server.Createobject("scripting.filesystemobject")

	IF instr(FunFileTarget,"/")>0 Then
		FunTargetArray = split(FunFileTarget,"/")
		for funfi=0 to ubound(FunTargetArray)-1
			If fordername&""="" then
				fordername = FunTargetArray(funfi) 			
			else
				fordername = fordername & "/" & FunTargetArray(funfi) 
			End IF
			if not FunFso.FolderExists( Server.MapPath(FunTargetStr&fordername) ) Then
				FunFso.CreateFolder( Server.MapPath(FunTargetStr&fordername) )
			end if
		Next
	End IF	
	
	'response.end
	if FunFso.fileexists(Server.MapPath(FunFileSource)) and not FunFso.fileexists(Server.MapPath(FunTargetStr&FunFileTarget)) then
		'response.write Server.MapPath(FunFileSource) & " to " & Server.MapPath(FunTargetStr&FunFileTarget)
		FunFso.MoveFile Server.MapPath(FunFileSource),Server.MapPath(FunTargetStr&FunFileTarget)
	ELSE
		response.write "fsoError:inc/func.asp"
		response.end
	end if
	Set FunFso = nothing
   '------------------------------------------------------------------------------
End Function

Function FunDelFile(FunFileSource)
	
   '------------------------------------------------------------------------------
   ' 判斷檔案是否已存在 by Mark
   ' 並且刪除檔案
   ' sample : call FunDelFile("temp/abc.txt")
   '------------------------------------------------------------------------------
	Set FunFso = Server.Createobject("scripting.filesystemobject")

	if FunFso.fileexists(Server.MapPath(FunFileSource)) then
		FunFso.Deletefile Server.MapPath(FunFileSource)
	end if
	Set FunFso = nothing
   '------------------------------------------------------------------------------
End Function


Function FunSplit(funarraystr,funsplitstr)
	IF funarraystr>"" Then
		FunSplit = split(funarraystr,funsplitstr)
	ELSE
		FunSplit = Array("")
	End IF
End Function


Function FnWriteItemStockLog(fnUnitNM,fnItemNo,fnInvType,fnSerialNo,fnQty,fnDescr)
	
	IF fnUnitNM&"">"" and fnItemNo&"">"" and fnInvType&"">"" and fnQty&"">"" and fnDescr&"">"" Then
		
		'寫入log start		
		fnLogId = Rdn12() & Rdn12()
		fnLogDate = dt2int(now())
		If cdbl(fnQty)>0 Then
			fnCredit = cdbl(fnQty)
			fnDebit = 0
			fnLogQty = fnCredit
			fnLogType = 1
		else
			fnDebit = cdbl(fnQty)*-1
			fnCredit = 0			
			fnLogQty = fnDebit			
			fnLogType = -1
		end if
	
		Set FunOBJconn = Server.CreateObject("ADODB.Connection")
		FunOBJconn.Open DBConnSTR
				
		Set FunRS = Server.CreateObject("ADODB.RecordSet")
		application.lock
			FunRS.open "Items_Stock_Log", FunOBJconn, 2, 3
			FunRS.addnew
			FunRS("LogId")=fnLogId
			FunRS("LogDate")=fnLogDate
			FunRS("LogType")=fnLogType			
			FunRS("UNITNM")=fnUnitNM
			FunRS("ItemNo")=fnItemNo
			FunRS("ItemInv")=fnInvType
			FunRS("SerialNo")=fnSerialNo
			FunRS("Credit")=fnCredit
			FunRS("Debit")=fnDebit
			FunRS("Qty")=fnLogQty
			FunRS("IsDone")="N"
			FunRS("Notes")=fnDescr
			FunRS("CREATE_DT")=now()
			FunRS("CREATE_USER")=Session("UID")
			
			FunRS.update
			FunRS.close
			
		application.unlock
		
		FnDoneLog = false	
		
		IF fnSerialNo&""="" Then
			'進行數量加減			
			FnSqlstr = "select * from Items_Stock where UnitNM='"&fnUnitNM&"' and ItemNo='"&fnItemNo&"'"
			FunRS.Open fnsqlstr,Objconn,3,2
				If FunRs.Recordcount > 0 Then
					FnUNIQID = FunRs("UNIQID")
					select case lcase(fnInvType)
						case "good"
							fnoldqty = cdbl(FunRs("qty"))
							fnFieldName = "qty"
						case "bad"
							fnoldqty = cdbl(FunRs("badqty"))					
							fnFieldName = "badqty"
					end select
					
					IF fnoldqty+cdbl(fnQty)>= 0 Then
						FunRs(fnFieldName)	= fnoldqty+cdbl(fnQty)
						FunRs("UPDATE_DT")	= now()
						FunRs("UPDATE_User") = Session("UID")
						FunRs.update
						FnDoneLog = true
					end if
				End IF
			FunRS.Close
		ELSE
			sqlstr = "exec sp_SysAutoCountSerialQty"
			FunOBJconn.Execute(sqlstr)
		End IF
		
		'完成Log
		IF FnDoneLog Then
			FnDoneLogStr = " UPDATE Items_Stock_Log set IsDone='Y',UPDATE_DT=getdate(),UPDATE_USER='"&Session("UID")&"' where LogId='"&fnLogId&"' "
			FunOBJconn.Execute(FnDoneLogStr)
		End IF
		
		FunOBJconn.Close
		Set FunRS = nothing
		Set FunOBJconn = nothing
	
	End IF

End Function


Function uSendHtmlMail(ToEmlAddress,FromEmlAddress,MailTitle,MailContext) 

	'寫入MailSendHistory
	MailID = Rdn12() & Rdn12()
	call WriteMailSendHistory(MailID,ToEmlAddress,FromEmlAddress,MailTitle,MailContext,"")
	
	
End Function 


'待發送郵件table
Function WriteMailSendHistory(FunMailID,ToEmlAddress,FromEmlAddress,MailTitle,MailContext,MailUrls)
	FunIsHttp = "N"
	FunIsSend = "N"
	IF MailUrls&"">"" Then
		FunIsHttp = "Y"
	end IF
	Set FunOBJconn = Server.CreateObject("ADODB.Connection")
	FunOBJconn.Open DBConnSTR
	Set FunRS = Server.CreateObject("ADODB.RecordSet")
	'寫入ptc_RecvStr
	Tablenm="MailSendHistory"
	application.lock
	
		FunRS.open ""& Tablenm &"", Objconn, 2, 3
		FunRS.addnew
		FunRS("MailID")=FunMailID
		FunRS("MailFrom")=FromEmlAddress
		FunRS("MailTo")=ToEmlAddress
		FunRS("MailSubject")=MailTitle
		IF MailContext&"">"" Then
			FunRS("MailBody")=MailContext
		End IF
		If MailUrls&"">"" Then
			FunRS("MailUrl")=MailUrls
		End IF
		FunRS("IsHttp")=FunIsHttp
		FunRS("IsSend")=FunIsSend
		FunRS("CREATE_DT")=thisDateTime()
				
	FunRS.update
	FunRS.close
	application.unlock		

	FunOBJconn.Close
	Set FunRS = nothing
	Set FunOBJconn = nothing

End Function


Function FnSQL(fnsqlstr,fndatatype)
	if trim(fnsqlstr&"")>"" Then
		select case fndatatype
			case "0" '字串
				fnsqlstr = replace(fnsqlstr,"'","''")
				FnSQL = " N'"& fnsqlstr & "' "
			case "1" '數字
				FnSQL = fnsqlstr
			case "2" '日期
				FnSQL = " '"& fnsqlstr & "' "	
			case "3" '系統日期函數 '
				'FnSQL = " geetdate() " 'MSSQL
				FnSQL = " now() " 'MySQL			
		end select 
	else
		FnSQL = " NULL "
	end if
end Function


Function FnWriteItemStockLog_New(fnUnitNM,fnItemNo,fnInvType,fnSerialNo,fnQty,fnDescr)
	
	IF fnUnitNM&"">"" and fnItemNo&"">"" and fnInvType&"">"" and fnQty&"">"" and fnDescr&"">"" Then
		
		'寫入log start		
		fnLogId = Rdn12() & Rdn12()
		fnLogDate = dt2int(now())
		If cdbl(fnQty)>0 Then
			fnCredit = cdbl(fnQty)
			fnDebit = 0
			fnLogQty = fnCredit
			fnLogType = 1
		else
			fnDebit = cdbl(fnQty)*-1
			fnCredit = 0			
			fnLogQty = fnDebit			
			fnLogType = -1
		end if
	
		Set FunOBJconn = Server.CreateObject("ADODB.Connection")
		FunOBJconn.Open DBConnSTR
		
		FunINSERTSQL = "INSERT INTO Items_Stock_Log (LogId " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ",LogDate " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ",LogType " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ",UNITNM " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ",ItemNo " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ",ItemInv " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ",SerialNo " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ",Credit " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ",Debit " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ",Qty " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ",IsDone " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ",Notes " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ",CREATE_DT " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ",CREATE_USER ) " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & " VALUES " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & "("&FnSQL(fnLogId,0)&" " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ","&FnSQL(fnLogDate,1)&" " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ","&FnSQL(fnLogType,1)&" " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ","&FnSQL(fnUnitNM,0)&" " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ","&FnSQL(fnItemNo,0)&" " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ","&FnSQL(fnInvType,0)&" " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ","&FnSQL(fnSerialNo,0)&" " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ","&FnSQL(fnCredit,1)&" " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ","&FnSQL(fnDebit,1)&" " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ","&FnSQL(fnLogQty,1)&" " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ",'N' " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ","&FnSQL(fnDescr,0)&" " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ",getdate() " & vbcrlf 
		FunINSERTSQL = FunINSERTSQL & ","&FnSQL(Session("UID"),0)&" ) " & vbcrlf 
		FunOBJconn.Execute(FunINSERTSQL)
		'response.write FunINSERTSQL & "<br>"
		FnDoneLog = false	
		
		IF fnSerialNo&""="" Then
			'進行數量加減			
			FnSqlstr = "select * from Items_Stock where UnitNM='"&fnUnitNM&"' and ItemNo='"&fnItemNo&"'"
			FunRS.Open fnsqlstr,Objconn,3,1
				If FunRs.Recordcount > 0 Then
					FnUNIQID = FunRs("UNIQID")
					select case lcase(fnInvType)
						case "good"
							fnoldqty = cdbl(FunRs("qty"))
							fnFieldName = "qty"
						case "bad"
							fnoldqty = cdbl(FunRs("badqty"))					
							fnFieldName = "badqty"
					end select
					
					IF fnoldqty+cdbl(fnQty)>= 0 Then
						FnUpdateStr = " UPDATE Items_Stock SET " & vbcrlf 
						FnUpdateStr = FnUpdateStr & " UPDATE_DT			=getdate() " & vbcrlf 
						FnUpdateStr = FnUpdateStr & " ,UPDATE_USER  	="&FnSQL(Session("UID"),0)&" " & vbcrlf 
						FnUpdateStr = FnUpdateStr & " ,"&fnFieldName&"  ="&fnoldqty+cdbl(fnQty)&" " & vbcrlf 
						FnUpdateStr = FnUpdateStr & " WHERE UNIQID  	="&FnUNIQID&" " & vbcrlf 
						FunOBJconn.Execute(FnUpdateStr)
						'response.write FnUpdateStr & "<br>"
						FnDoneLog = true
					end if
				End IF
			FunRS.Close
		ELSE
			sqlstr = "exec sp_SysAutoCountSerialQty"
			FunOBJconn.Execute(sqlstr)
		End IF
		
		'完成Log
		IF FnDoneLog Then
			FnDoneLogStr = " UPDATE Items_Stock_Log set IsDone='Y',UPDATE_DT=getdate(),UPDATE_USER='"&Session("UID")&"' where LogId='"&fnLogId&"' "
			FunOBJconn.Execute(FnDoneLogStr)
		End IF
		
		FunOBJconn.Close
		Set FunRS = nothing
		Set FunOBJconn = nothing
	
	End IF

End Function



Function fnCdbl(fnNumber)
	if My_IsNumeric(fnNumber) then
		fnCdbl = cdbl(fnNumber)
	else
		fnCdbl = 0
	end if
End Function


Function My_IsNumeric(fnvalue)
	If trim(fnvalue&"")>"" then
		My_IsNumeric = IsNumeric(trim(fnvalue&""))
	else
		My_IsNumeric = false
	end if  
End Function



function fnCheckItem(fncatego,fnfieldname,fnfieldvalue)

	Set FunOBJconn = Server.CreateObject("ADODB.Connection")
	FunOBJconn.Open DBConnSTR

	fnStr = ""
	sqlstr = "select item_txt,item_val,item_note_yn,item_note_html,item_breakrow from FORMITEM where item_type='"&fncatego&"' order by item_sort "
	set FunRs = FunOBJconn.execute(sqlstr)
	set fnRadio_ARY = nothing
	if not FunRs.eof then
		fnRadio_ARY = FunRs.getrows()
	end if
	FunRs.close	
	if IsArray(fnRadio_ARY) then
		for rows=0 to ubound(fnRadio_ARY,2)
			rdoid = fnfieldname&"_"&rows

			fnvalue_ARY = FunSplit(fnfieldvalue&"",",")

			ischk = false
			If IsArray(fnvalue_ARY) then
				for xn=0 to ubound(fnvalue_ARY)				
					if trim(fnRadio_ARY(1,rows)) = trim(fnvalue_ARY(xn)) then
						ischk = true
						exit for
					end if
				next
			end if


			if trim(fnRadio_ARY(4,rows))="Y" then
				fnStr = fnstr & "<div>"
			else
				fnStr = fnstr & "<span>"
			end if

			fnStr = fnstr & "<input type=""checkbox"" name="""&fnfieldname&""" value="""&trim(fnRadio_ARY(1,rows))&""" "
			if ischk then
				fnStr = fnstr & " checked "
			end if
			fnStr = fnstr & " id="""&rdoid&"""><label for="""&rdoid&""" data-html="""&trim(fnRadio_ARY(0,rows))&""">"&trim(fnRadio_ARY(0,rows))&"</label> "

			if trim(fnRadio_ARY(2,rows))="Y" then
				fnStr = fnstr &  trim(fnRadio_ARY(3,rows))
			end if

			if trim(fnRadio_ARY(4,rows))="Y" then
				fnStr = fnstr & "</div>" & vbcrlf
			else
				fnStr = fnstr & "</span>" & vbcrlf
			end if

		next
	end if

	FunOBJconn.Close
	Set FunRS = nothing
	Set FunOBJconn = nothing

	fnCheckItem = fnStr
end function


function fnRadioItem(fncatego,fnfieldname,fnfieldvalue)

	Set FunOBJconn = Server.CreateObject("ADODB.Connection")
	FunOBJconn.Open DBConnSTR

	fnStr = ""
	sqlstr = "select item_val+'-'+item_txt,item_val,item_note_yn,item_note_html,item_breakrow from FORMITEM where item_type='"&fncatego&"' order by item_sort "
	set FunRs = FunOBJconn.execute(sqlstr)
	set fnRadio_ARY = nothing
	if not FunRs.eof then
		fnRadio_ARY = FunRs.getrows()
	end if
	FunRs.close	
	if IsArray(fnRadio_ARY) then
		fnStr = "<span col-value="""">"
		for rows=0 to ubound(fnRadio_ARY,2)
			rdoid = fnfieldname&"_"&rows
			ischk = false
			if trim(fnRadio_ARY(1,rows)) = trim(fnfieldvalue) then
				ischk = true
			end if

			if trim(fnRadio_ARY(4,rows))="Y" then
				fnStr = fnstr & "<div>"
			else
				fnStr = fnstr & "<span>"
			end if

			fnStr = fnstr & "<input type=""radio"" name="""&fnfieldname&""" value="""&trim(fnRadio_ARY(1,rows))&""" col-type=""radio"" "
			if ischk then
				fnStr = fnstr & " checked "
			end if
			fnStr = fnstr & " id="""&rdoid&"""><label for="""&rdoid&""" data-html="""&trim(fnRadio_ARY(0,rows))&""">"&trim(fnRadio_ARY(0,rows))&"</label> "

			if trim(fnRadio_ARY(2,rows))="Y" then
				fnStr = fnstr &  trim(fnRadio_ARY(3,rows))
			end if

			if trim(fnRadio_ARY(4,rows))="Y" then
				fnStr = fnstr & "</div>" & vbcrlf
			else
				fnStr = fnstr & "</span>" & vbcrlf
			end if

		next
		fnStr = fnstr & "</span>" & vbcrlf
	end if

	FunOBJconn.Close
	Set FunRS = nothing
	Set FunOBJconn = nothing

	fnRadioItem = fnStr
end function


function fnComboItem(fncatego,fnfieldname,fnfieldvalue)

	Set FunOBJconn = Server.CreateObject("ADODB.Connection")
	FunOBJconn.Open DBConnSTR

	fnStr = ""
	sqlstr = "select item_val+'.'+item_txt,item_val,item_note_yn,item_note_html,item_breakrow from FORMITEM where item_type='"&fncatego&"' order by item_sort "
	set FunRs = FunOBJconn.execute(sqlstr)
	set fnRadio_ARY = nothing
	if not FunRs.eof then
		fnCombo_ARY = FunRs.getrows()
	end if
	FunRs.close	
	if IsArray(fnCombo_ARY) then
		fnStr = "<select name="""&fnfieldname&""" onchange=""doselect(this);"">" & vbcrlf
		fnStr = fnStr & "<option value="""" data-YN=""N"" data-afterhtml=""""></option>" & vbcrlf
		for rows=0 to ubound(fnCombo_ARY,2)
			rdoid = fnfieldname&"_"&rows
			ischk = false
			if trim(fnCombo_ARY(1,rows)) = trim(fnfieldvalue) then
				ischk = true
			end if

			fnStr = fnStr & "<option value="""&trim(fnCombo_ARY(1,rows)) &""" data-YN="""&trim(fnCombo_ARY(2,rows))&""" data-afterhtml="""&replace(trim(fnCombo_ARY(3,rows)&""),"""","'")&""" "
			if ischk then
			fnStr = fnStr & " selected "
			end if
			fnStr = fnStr & ">"&trim(fnCombo_ARY(0,rows)) &"</option>" & vbcrlf
		next
		fnStr = fnStr & "</select>" & vbcrlf
		fnStr = fnStr & "<span col=""af_span""></span>" & vbcrlf
	end if

	FunOBJconn.Close
	Set FunRS = nothing
	Set FunOBJconn = nothing


	fnComboItem = fnStr
end function


function fnDatetime(fnfieldname,fndate,fntimestr)
	fnstr = ""
	fnstr = fnstr & "日期：<input type=""text"" name="""&fnfieldname&"_DT"" datarole=""pickdate"" style=""width:80px;"" value="""&fndate&""" maxlength=""10"">(YYYY/MM/DD)&nbsp;&nbsp;"
	fnstr = fnstr & "時間：<input type=""text"" name="""&fnfieldname&"_HH"" style=""width:40px;"" value="""&left(fntimestr,2)&"""  maxlength=""2"" col=""HH"" >:<input type=""text"" name="""&fnfieldname&"_MM"" style=""width:40px;"" value="""&right(fntimestr,2)&""" maxlength=""2"" col=""MM"">(HH:MM)"
	fnDatetime = fnstr
end function


function fntime(fnfieldname,fntimestr)
	fnstr = ""
	fnstr = fnstr & "<input type=""text"" name="""&fnfieldname&"_HH"" style=""width:40px;"" value="""&left(fntimestr,2)&"""  maxlength=""2"" col=""HH"" >:<input type=""text"" name="""&fnfieldname&"_MM"" style=""width:40px;"" value="""&right(fntimestr,2)&""" maxlength=""2"" col=""MM"">(HH:MM)"
	fntime = fnstr
end function


function fnCheckItem_Single(fncatego,fnfieldname,fnfieldvalue)

	Set FunOBJconn = Server.CreateObject("ADODB.Connection")
	FunOBJconn.Open DBConnSTR

	fnStr = "<span>"
	sqlstr = "select item_txt,item_val,item_note_yn,item_note_html,item_breakrow from FORMITEM where item_type='"&fncatego&"' order by item_sort "
	set FunRs = FunOBJconn.execute(sqlstr)
	set fnRadio_ARY = nothing
	if not FunRs.eof then
		fnRadio_ARY = FunRs.getrows()
	end if
	FunRs.close	
	if IsArray(fnRadio_ARY) then
		for rows=0 to ubound(fnRadio_ARY,2)
			rdoid = fnfieldname&"_"&rows

			ischk = false
			if trim(fnRadio_ARY(1,rows)) = trim(fnfieldvalue) then
				ischk = true
			end if

			if trim(fnRadio_ARY(4,rows))="Y" then
				fnStr = fnstr & "<div>"
			else
				fnStr = fnstr & "<span>"
			end if

			fnStr = fnstr & "<input type=""checkbox"" name="""&fnfieldname&""" value="""&trim(fnRadio_ARY(1,rows))&""" col-type=""single"" "
			if ischk then
				fnStr = fnstr & " checked "
			end if
			fnStr = fnstr & " id="""&rdoid&"""><label for="""&rdoid&""" data-html="""&trim(fnRadio_ARY(0,rows))&""">"&trim(fnRadio_ARY(0,rows))&"</label> "

			if trim(fnRadio_ARY(2,rows))="Y" then
				fnStr = fnstr &  trim(fnRadio_ARY(3,rows))
			end if

			if trim(fnRadio_ARY(4,rows))="Y" then
				fnStr = fnstr & "</div>" & vbcrlf
			else
				fnStr = fnstr & "</span>" & vbcrlf
			end if

		next
	end if
	
	fnStr = fnstr & "</span>" & vbcrlf
	FunOBJconn.Close
	Set FunRS = nothing
	Set FunOBJconn = nothing

	fnCheckItem_Single = fnStr
end function



Function fnArraySort(fnary,sorttype)
	Dim fni
	if IsArray(fnary) then
		select case sorttype
			case "asc"





			case "desc"

				fnmax=ubound(fnary)
				For fni=0 to fnmax  
				   For fnj=fni to fnmax  
				      if fnary(fni)<fnary(fnj) then 
				          TemporalVariable=fnary(fni) 
				          fnary(fni)=fnary(fnj) 
				          fnary(fnj)=TemporalVariable 
				     end if 
				   next  
				next 

		end select
		fnArraySort = fnary
	end if
end Function



Function FunCheckFile(FunFileTarget)

   '------------------------------------------------------------------------------
   ' 判斷檔案是否已存在 by Mark
   ' 並且複製檔案至特定目錄
   ' sample : call FunCopyFile("temp/abc.txt","objtect/txt/abc.txt")
   '------------------------------------------------------------------------------

	'判斷複製目錄是否存在，不存在則建立之
	IF left(FunFileTarget,1)="/" Then
		FunTargetStr = "/"
		FunFileTarget = right(FunFileTarget,len(FunFileTarget)-1)
	End IF

	Set FunFso = Server.Createobject("scripting.filesystemobject")

	IF instr(FunFileTarget,"/")>0 Then
		FunTargetArray = split(FunFileTarget,"/")
		for funfi=0 to ubound(FunTargetArray)-1
			If fordername&""="" then
				fordername = FunTargetArray(funfi)
			else
				fordername = fordername & "/" & FunTargetArray(funfi)
			End IF
			if not FunFso.FolderExists( Server.MapPath(FunTargetStr&fordername) ) Then
				FunFso.CreateFolder( Server.MapPath(FunTargetStr&fordername) )
			end if
		Next
	End IF

	'response.end
	if FunFso.fileexists(Server.MapPath(FunTargetStr&FunFileTarget)) then
		FunCheckFile = true
	ELSE
		FunCheckFile = false
	end if
	Set FunFso = nothing
   '------------------------------------------------------------------------------
End Function


%>
