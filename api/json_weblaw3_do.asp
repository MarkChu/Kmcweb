<%@Language=VBScript codepage=65001 %>
<%Response.CodePage = 65001%>
<%Response.Charset = "utf-8"%>
<%Response.ContentType= "application/json"%>
<!--#include file="json_common.asp"-->
<!--#include file="../inc/common.asp"-->
<!--#include file="../inc/md5.asp"-->
<!--#include file="clsUpload.asp"-->
<!--#include file="../inc/Func.asp"-->
<!--#include file="../inc/JSON_2.0.4.asp"-->

<%
'on error resume next
'Server.ScriptTimeout = 900

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
status_desc = ""


Dim Upload
Dim FileName
Dim Folder


Set Upload = New clsUpload

Folder = Server.MapPath("../weblaw") & "\"

'Response.Write Upload.DebugText
act = Upload.Fields("act").Value
uniqid 	= Upload.Fields("uniqid").Value
lawid 	= Upload.Fields("lawid").Value
lawcatego = Upload.Fields("lawcatego").Value
title 		= Upload.Fields("title").Value
sortid 		= Upload.Fields("sortid").Value
ison 		= Upload.Fields("ison").Value
url 	= Upload.Fields("url").Value


' Release upload object from memory
Set Upload = Nothing




if status="0000" then

	select case  act
		case "add"
			sqlstr = "insert into weblaw (lawcatego " & vbcrlf 
			sqlstr = sqlstr & ",lawid " & vbcrlf 
			sqlstr = sqlstr & ",lawtitle " & vbcrlf 
			sqlstr = sqlstr & ",url " & vbcrlf 
			sqlstr = sqlstr & ",ison " & vbcrlf 
			sqlstr = sqlstr & ",sortid " & vbcrlf 
			sqlstr = sqlstr & ",createdate " & vbcrlf 
			sqlstr = sqlstr & ",createuser " & vbcrlf 
			sqlstr = sqlstr & ") values ("&FnSQL(lawcatego,1)&" " & vbcrlf 
			sqlstr = sqlstr & ","&FnSQL(lawid,0)&" " & vbcrlf 
			sqlstr = sqlstr & ","&FnSQL(title,0)&" " & vbcrlf 
			sqlstr = sqlstr & ","&FnSQL(url,0)&" " & vbcrlf 
			sqlstr = sqlstr & ","&FnSQL(ison,0)&" " & vbcrlf 
			sqlstr = sqlstr & ","&FnSQL(sortid,1)&" " & vbcrlf 
			sqlstr = sqlstr & ",getdate() " & vbcrlf 
			sqlstr = sqlstr & ","&FnSQL(session("userid"),0)&" " & vbcrlf 
			sqlstr = sqlstr & "); " & vbcrlf 
			objconn.execute(sqlstr)

		case "edit"

			sqlstr = "update weblaw set lawtitle = "&FnSQL(title,0)&" " & vbcrlf 
			sqlstr = sqlstr & ",url = "&FnSQL(url,0)&" " & vbcrlf 
			sqlstr = sqlstr & ",sortid = "&FnSQL(sortid,1)&" " & vbcrlf 
			sqlstr = sqlstr & ",ison = "&FnSQL(ison,0)&" " & vbcrlf 
			sqlstr = sqlstr & "where lawid='"&lawid&"' " & vbcrlf 
			objconn.execute(sqlstr)


	end select


end if

Dim jsa
Set jsa = jsObject()

jsa("status") = status
jsa("status_desc") = status_desc

jsa.Flush
response.end%>