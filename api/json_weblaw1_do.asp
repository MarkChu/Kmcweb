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
lawcontent 	= Upload.Fields("lawcontent").Value


detid_str = Upload.Fields("detid").Value
isdel_str = Upload.Fields("isdel").Value
chcatego_str = Upload.Fields("chcatego").Value
chtitle_str = Upload.Fields("chtitle").Value
chcontent_str = Upload.Fields("chcontent").Value
detsortid_str = Upload.Fields("detsortid").Value

splitstr = "@kmc@"

detid_Array = FunSplit(detid_str,splitstr)
isdel_Array = FunSplit(isdel_str,splitstr)
chcatego_Array = FunSplit(chcatego_str,splitstr)
chtitle_Array = FunSplit(chtitle_str,splitstr)
chcontent_Array = FunSplit(chcontent_str,splitstr)
detsortid_Array = FunSplit(detsortid_str,splitstr)

' Release upload object from memory
Set Upload = Nothing




if status="0000" then

	select case  act
		case "add"
			sqlstr = "insert into weblaw (lawcatego " & vbcrlf 
			sqlstr = sqlstr & ",lawid " & vbcrlf 
			sqlstr = sqlstr & ",lawtitle " & vbcrlf 
			sqlstr = sqlstr & ",lawcontent " & vbcrlf 
			sqlstr = sqlstr & ",ison " & vbcrlf 
			sqlstr = sqlstr & ",sortid " & vbcrlf 
			sqlstr = sqlstr & ",createdate " & vbcrlf 
			sqlstr = sqlstr & ",createuser " & vbcrlf 
			sqlstr = sqlstr & ") values ("&FnSQL(lawcatego,1)&" " & vbcrlf 
			sqlstr = sqlstr & ","&FnSQL(lawid,0)&" " & vbcrlf 
			sqlstr = sqlstr & ","&FnSQL(title,0)&" " & vbcrlf 
			sqlstr = sqlstr & ","&FnSQL(lawcontent,0)&" " & vbcrlf 
			sqlstr = sqlstr & ","&FnSQL(ison,0)&" " & vbcrlf 
			sqlstr = sqlstr & ","&FnSQL(sortid,1)&" " & vbcrlf 
			sqlstr = sqlstr & ",getdate() " & vbcrlf 
			sqlstr = sqlstr & ","&FnSQL(session("userid"),0)&" " & vbcrlf 
			sqlstr = sqlstr & "); " & vbcrlf 
			objconn.execute(sqlstr)

		case "edit"

			sqlstr = "update weblaw set lawtitle = "&FnSQL(title,0)&" " & vbcrlf 
			sqlstr = sqlstr & ",lawcontent = "&FnSQL(lawcontent,0)&" " & vbcrlf 
			sqlstr = sqlstr & ",sortid = "&FnSQL(sortid,1)&" " & vbcrlf 
			sqlstr = sqlstr & ",ison = "&FnSQL(ison,0)&" " & vbcrlf 
			sqlstr = sqlstr & "where lawid='"&lawid&"' " & vbcrlf 
			objconn.execute(sqlstr)


	end select


	If IsArray(detid_Array) then
		for r=0 to ubound(detid_Array)
			detact = "add"
			if IsNumeric(detid_Array(r)) then
				if cdbl(detid_Array(r))>0 then
					detact = "edit"
				end if
			end if

			if isdel_Array(r)="Y" then
				detact = "del"
			end if

			select case detact
				case "add"

					sqlstr = "insert into weblawdet (lawid " & vbcrlf 
					sqlstr = sqlstr & ",chcatego " & vbcrlf 
					sqlstr = sqlstr & ",chtitle " & vbcrlf 
					sqlstr = sqlstr & ",chcontent " & vbcrlf 
					sqlstr = sqlstr & ",sortid " & vbcrlf 
					sqlstr = sqlstr & ") values ("&FnSQL(lawid,0)&" " & vbcrlf 
					sqlstr = sqlstr & ","&FnSQL(chcatego_Array(r),0)&" " & vbcrlf 
					sqlstr = sqlstr & ","&FnSQL(chtitle_Array(r),0)&" " & vbcrlf 
					sqlstr = sqlstr & ","&FnSQL(chcontent_Array(r),0)&" " & vbcrlf 
					sqlstr = sqlstr & ","&FnSQL(detsortid_Array(r),1)&" " & vbcrlf 
					sqlstr = sqlstr & ") " & vbcrlf 
					objconn.execute(sqlstr)

				case "edit"

					sqlstr = "update weblawdet set chcatego = "&FnSQL(chcatego_Array(r),0)&" " & vbcrlf 
					sqlstr = sqlstr & ",chtitle = "&FnSQL(chtitle_Array(r),0)&" " & vbcrlf 
					sqlstr = sqlstr & ",chcontent = "&FnSQL(chcontent_Array(r),0)&" " & vbcrlf 
					sqlstr = sqlstr & ",sortid="&FnSQL(detsortid_Array(r),1)&" " & vbcrlf 
					sqlstr = sqlstr & "where uniqid="&detid_Array(r)&" " & vbcrlf 
					objconn.execute(sqlstr)


				case "del"

					sqlstr = "delete from weblawdet where uniqid="&detid_Array(r)&" " & vbcrlf 
					objconn.execute(sqlstr)

			end select


		next
	end if


end if

Dim jsa
Set jsa = jsObject()

jsa("status") = status
jsa("status_desc") = status_desc

jsa.Flush
response.end%>