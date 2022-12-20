<!-- #include file="common.asp" -->
<%
'toMail
ToEmail = "markchu929@gmail.com"
'from
FromEmail = "elaine@nec.com.tw"
'subject
Mailtitle = "test mail"
'content
MailContent = "content mail"

'uSendHtmlMail(ToEmlAddress,FromEmlAddress,MailTitle,MailContext) 
call uSendHtmlMail(ToEmail,FromEmail,Mailtitle,MailContent)

response.write "發送至" & ToEmail & ",發送成功!"

%>
<%

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
		FunRS("CREATE_DT")=now()
				
	FunRS.update
	FunRS.close
	application.unlock		

	FunOBJconn.Close
	Set FunRS = nothing
	Set FunOBJconn = nothing

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
	For FunI = 1 To 12
	Session("Temp_" & FunI) = Clng(Asc(Mid(Code,FunI,1)))
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
	For FunI = 1 To 12
		If Session("Code_" & FunI) > 9 Then
			Session("Code_" & FunI) = Chr(Session("Code_" & FunI) + 55)
		End If
		Password = Password & Session("Code_" & FunI)
	Next
	
	Rdn12 = Password
	
End Function



%>