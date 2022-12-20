<%Response.CodePage = 65001%>
<%Response.Charset = "utf-8"%>
<%Response.ContentType= "application/json"%>
<!--#include file="json_common.asp"-->
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
uniqid = request("uniqid")

if uniqid>"" then
	uniqid_ary = split(uniqid,",")
	if isArray(uniqid_ary) then
		sql_uniqid = "''"
		for r=0 to ubound(uniqid_ary)
			sql_uniqid = sql_uniqid & ",'"&trim(uniqid_ary(r))&"'"
		next
	end if
else
	status = "9999"
	status_desc = "請至少勾選一個項目。"
end if




if status="0000" then

	set sql_cmd = Server.CreateObject("ADODB.Command") 
	sql_cmd.ActiveConnection = Objconn
	sql_cmd.CommandText = "delete from sms where unique_id in ("&sql_uniqid&") " 
	'ADO.CreateParameter(name,type,direction,size,value)
	set rs = sql_cmd.Execute
end if

Dim jsa
Set jsa = jsObject()

jsa("status") = status
jsa("status_desc") = status_desc

jsa.Flush
response.end%>