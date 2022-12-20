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

Folder = Server.MapPath("../accounts") & "\"

'Response.Write Upload.DebugText
act = Upload.Fields("act").Value
uniqid = Upload.Fields("uniqid").Value
postdate 	= Upload.Fields("postdate").Value
title 		= Upload.Fields("title").Value
catego 		= Upload.Fields("catego").Value
ison 		= Upload.Fields("ison").Value

' Grab the file name
set theFile = Upload.Fields("File1")
FileSize = theFile.Length
'

if FileSize>0 then
	OriFileName = theFile.FileName	
	Ext_ary = split(OriFileName,".")
	NewFileName = md5(OriFileName&now)&"."&Ext_ary(1)
	' Get path to save file to
	' Save the binary data to the file system
	theFile.SaveAs Folder & NewFileName

end if

' Release upload object from memory
Set Upload = Nothing


if status="0000" then

	select case  act
		case "add"
			sqlstr = "insert into accountbbs (catego " & vbcrlf 
			sqlstr = sqlstr & ",title " & vbcrlf 
			sqlstr = sqlstr & ",postdate " & vbcrlf 
			sqlstr = sqlstr & ",ison " & vbcrlf 
			sqlstr = sqlstr & ",attachfile1 " & vbcrlf 
			sqlstr = sqlstr & ",attachname1 " & vbcrlf 
			sqlstr = sqlstr & ",createdate " & vbcrlf 
			sqlstr = sqlstr & ",createuser " & vbcrlf 
			sqlstr = sqlstr & ") values ("&FnSQL(catego,0)&" " & vbcrlf 
			sqlstr = sqlstr & ","&FnSQL(title,0)&" " & vbcrlf 
			sqlstr = sqlstr & ","&FnSQL(postdate,2)&" " & vbcrlf 
			sqlstr = sqlstr & ","&FnSQL(ison,0)&" " & vbcrlf 
			sqlstr = sqlstr & ","&FnSQL(NewFileName,0)&" " & vbcrlf 
			sqlstr = sqlstr & ","&FnSQL(OriFileName,0)&" " & vbcrlf 
			sqlstr = sqlstr & ",getdate() " & vbcrlf 
			sqlstr = sqlstr & ","&FnSQL(session("userid"),0)&" " & vbcrlf 
			sqlstr = sqlstr & ") " & vbcrlf 
			objconn.execute(sqlstr)
		case "edit"


			sqlstr = "update accountbbs set catego = "&FnSQL(catego,0)&" " & vbcrlf 
			sqlstr = sqlstr & ",title = "&FnSQL(title,0)&" " & vbcrlf 
			sqlstr = sqlstr & ",postdate = "&FnSQL(postdate,2)&" " & vbcrlf 
			sqlstr = sqlstr & ",ison = "&FnSQL(ison,0)&" " & vbcrlf 
			if OriFileName>"" then
				sqlstr = sqlstr & ",attachfile1="&FnSQL(NewFileName,0)&" " & vbcrlf 
				sqlstr = sqlstr & ",attachname1="&FnSQL(OriFileName,0)&" " & vbcrlf 
			end if
			sqlstr = sqlstr & "where uniqid="&uniqid&" " & vbcrlf 
			objconn.execute(sqlstr)


	end select


end if

Dim jsa
Set jsa = jsObject()

jsa("status") = status
jsa("status_desc") = status_desc

jsa.Flush
response.end%>