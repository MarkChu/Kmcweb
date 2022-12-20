<!--#include file="Common.asp"-->
<!--#include file="Func.asp"-->
<%
Response.Charset="utf-8"
uniqid = request("uniqid")




'此頁面為取得人員之相關資料
SQLSTR = "SELECT * FROM LICENCE where UNIQID="&uniqid&" "
rs.Open SQLSTR,Objconn,3,1
	IF rs.recordcount > 0 Then
		LICE_STR1 = rs("LICE_STR1")
		LICE_STR3 = rs("LICE_STR2")
	ELSE
		ErrorSTR = "無此證照資料!!"
	End IF
rs.Close

'返回之資料值
response.write "#@#LICE_STR1#:#" & LICE_STR1
response.write "#@#LICE_STR3#:#" & LICE_STR3
response.write "#@#ERROR#:#" & ErrorSTR
%>
<!--#include file="Close_Connection.asp"-->

