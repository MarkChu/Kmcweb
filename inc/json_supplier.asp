<%Response.CodePage = 65001%>
<%Response.Charset = "utf-8"%>
<%Response.ContentType= "application/json"%>
<!--#include file="common.asp"-->
<!--#include file="JSON_2.0.4.asp"-->
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

'response.write sqlstr
'response.end

datarows = 0
'json start
Dim rs, jsa
c1_total = 0

sqlstr = " select A01 as supplierid,NAME as suppliername  " & vbcrlf 
sqlstr = sqlstr & " from gw_aa " & vbcrlf 
sqlstr = sqlstr & " where 1=1 " & vbcrlf 
if request("p")&"">"" then
sqlstr = sqlstr & " and (A01 like '"&request("p")&"%' or NAME like '%"&request("p")&"%') "
end if
sqlstr = sqlstr & " order by 1 " & vbcrlf 
Set rs = Objconn.Execute(SQLSTR)
if not rs.eof then
	lv1ARY = rs.getrows()
end if
rs.close
Set jsa = jsArray()


if IsArray(lv1ARY) then
	for lv1=0 to ubound(lv1ARY,2)
		Set jsa(Null) = jsObject()
		c1_id = lv1ARY(0,lv1)
		c1_name = lv1ARY(1,lv1) 

		jsa(Null)("supplierid") = trim(cstr(c1_id))
		jsa(Null)("suppliername") = trim(cstr(c1_name))

	next
end if

jsa.Flush
response.end%>