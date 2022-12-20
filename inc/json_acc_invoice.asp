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

yymmdd = request("yymmdd")
invtype = request("invtype")

select case right(left(yymmdd,6),2)
	case "01","02"
		acc_MM = "02"
	case "03","04"
		acc_MM = "04"
	case "05","06"
		acc_MM = "06"
	case "07","08"
		acc_MM = "08"
	case "09","10"
		acc_MM = "10"
	case "11","12"
		acc_MM = "12"		
end select

acc_YYMM = left(yymmdd,4)&acc_MM


sqlstr = " select invoiceno,startno,SN,uniqid " & vbcrlf 
sqlstr = sqlstr & " from acc_invoice  " & vbcrlf 
sqlstr = sqlstr & " where yymm='"&acc_YYMM&"' " & vbcrlf 
sqlstr = sqlstr & " and invoicetype='"&invtype&"' " & vbcrlf 
sqlstr = sqlstr & " order by invoiceno,startno " & vbcrlf 
Set rs = Objconn.Execute(SQLSTR)
if not rs.eof then
	lv1ARY = rs.getrows()
end if
rs.close
Set jsa = jsArray()


if IsArray(lv1ARY) then
	for lv1=0 to ubound(lv1ARY,2)
		Set jsa(Null) = jsObject()
		invoiceno = lv1ARY(0,lv1)
		startno = lv1ARY(1,lv1) 
		SN = cdbl("0"&lv1ARY(2,lv1)) 
		uniqid = lv1ARY(3,lv1)
		invoicestr = invoiceno & right("00000000"&(cdbl(startno)+SN),8)

		jsa(Null)("invoiceno") = trim(cstr(invoicestr))
		jsa(Null)("uniqid") = trim(cstr(uniqid))

	next
end if

jsa.Flush
response.end%>