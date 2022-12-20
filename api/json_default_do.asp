<%Response.CodePage = 65001%>
<%Response.Charset = "utf-8"%>
<%Response.ContentType= "application/json"%>
<!--#include file="../inc/common.asp"-->
<!--#include file="../inc/JSON_2.0.4.asp"-->
<%
Function My_IsNumeric(value)
    My_IsNumeric = False
    If IsNull(value) Then Exit Function
    My_IsNumeric = IsNumeric(CStr(value))
End Function

Function checkhack(keystr)
	allcheckstr = "'%<>-/*;"
	checkstr = trim(keystr)
	for Funi=1 to len(allcheckstr)
		checkstr = replace(checkstr,mid(allcheckstr,Funi,1),"")
	Next
	checkhack = checkstr
End Function

status = "0000"
id=request("a")
pw=request("b")
auth = request("c")

if auth<>session("KmcAuthCode") then
	status = "9999"
	status_desc = "驗證碼輸入錯誤!!"
end if

if status="0000" then

	set sql_cmd = Server.CreateObject("ADODB.Command") 
	sql_cmd.ActiveConnection = Objconn
	sql_cmd.CommandText = "select [name],[account_id],[dept] from members where account_id = ? and account_password = ? " 
	'ADO.CreateParameter(name,type,direction,size,value)
	sql_cmd.Parameters.Append sql_cmd.CreateParameter("account_id",202,1,20,id)
	sql_cmd.Parameters.Append sql_cmd.CreateParameter("password",202,1,20,pw)
	set rs = sql_cmd.Execute
	if not rs.eof then
		userid = rs(1)

		sqlstr = " set NOCOUNT on;set ANSI_WARNINGS on; set ANSI_NULLS on; " & vbcrlf 	
		sqlstr = sqlstr & " exec spSyncData "
		set rs2 = Objconn.execute(sqlstr)
		if rs2.eof then
			status = "9999"
			status_desc = "資料同步失敗!!"			
		else
			sqlstr = "select [name],[account_id],[dept],webaccount,sms,internet,weblaw from members where account_id='"&userid&"'"
			set rs3 = objconn.execute(sqlstr)
			if not rs3.eof then
				session("username")=rs3(0)
				session("userid")=rs3(1)
				session("dept")=rs3(2)
				session("account")=cdbl("0"&rs3(3))
				session("sms")=cdbl("0"&rs3(4))
				session("internet")=cdbl("0"&rs3(5))
				session("weblaw")=cdbl("0"&rs3(6))
			else
				status = "9999"
				status_desc = "帳號驗證錯誤!!"
			end if
			rs3.close
		end if
		rs2.close
	else
		status = "9999"
		status_desc = "帳號驗證錯誤!!"
	end if
	rs.close

end if

Dim jsa
Set jsa = jsObject()

jsa("status") = status
jsa("status_desc") = status_desc

jsa.Flush
response.end%>