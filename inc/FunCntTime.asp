<!--#include file="Common.asp"-->
<!--#include file="Func.asp"-->
<%
Response.Charset="utf-8"
tm1 = Request("tm1")
tm2 = Request("tm2")

IF not isNumeric(tm1) or not isNumeric(tm2) Then
	ErrorSTR = "時間格式錯誤!!請確認時間格式為HHMM,例如15:10則輸入1510。"
end IF

dt1 = int2dt(request("dt1"))
dt2 = int2dt(request("dt2"))
tm1 = CDATE(dt1&" "&left(tm1,2) & ":" & right(tm1,2))
tm2 = CDATE(dt2&" "&left(tm2,2) & ":" & right(tm2,2))

Minutes = datediff("n",tm1,tm2)

hours = round(Minutes/60,2)


'返回之資料值
response.write "#@#HOURS#:#" & hours
response.write "#@#ERROR#:#" & ErrorSTR
%>
<!--#include file="Close_Connection.asp"-->

