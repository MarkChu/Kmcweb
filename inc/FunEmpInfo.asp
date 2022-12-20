<!--#include file="Common.asp"-->
<!--#include file="Func.asp"-->
<%
Response.Charset="utf-8"
empid = request("empid")

AuthSQL = " AND UnitNM in ("&Session("AuthUnitNM")&") "






'此頁面為取得人員之相關資料
SQLSTR = "SELECT * FROM Employee where EmpID='"&empid&"' " & AuthSQL
rs.Open SQLSTR,Objconn,3,1
	IF rs.recordcount > 0 Then
		EmpID = rs("EmpID")
		EmpNM = rs("EmpNM")
		UNITNM = rs("UNITNM")
	ELSE
		ErrorSTR = "您無此人員資料修改的權限 或 無此人員資料!!"
	End IF
rs.Close


IF request("train_uniqid")&"">"" Then

	TRAIN_AUTH = cint(GetFieldValue("TRAINING","UNIQID="&request("train_uniqid"),"TRAIN_AUTH"))
	IF TRAIN_AUTH = 1 Then
		TRAIN_AUTH_UNITNM = "'"&replace(GetFieldValue("TRAINING","UNIQID="&request("train_uniqid"),"TRAIN_AUTH_UNITNM"),",","','")&"'"
		IF cint(GetFieldValue("employee","empid='"&EMPID&"' and unitnm in ("&TRAIN_AUTH_UNITNM&")","count(*)"))=0 Then
			ErrorSTR = ErrorSTR & vbcrlf & "此繼續教育並未對此人員之單位開放，請重新選擇人員!!"
		End IF
	End IF
	
	LICE_UNIQID = GetFieldValue("TRAINING","UNIQID="&request("train_uniqid"),"LICE_UNIQID")
	IF cint(GetFieldValue("employee_licence","on_Flag in ('Y','N') and empid='"&EMPID&"' and Lice_UniqID="&LICE_UNIQID,"count(*)"))=0 Then
		ErrorSTR = ErrorSTR & vbcrlf & "此人員並無取得該繼續教育之證照，請重新確認!!"
	End IF

End IF







'返回之資料值
response.write "#@#EMPID#:#" & EmpID
response.write "#@#EMPNM#:#" & EmpNM
response.write "#@#UNITNM#:#" & UNITNM
response.write "#@#ERROR#:#" & ErrorSTR
%>
<!--#include file="Close_Connection.asp"-->

