<%Response.CodePage = 65001%>
<%Response.Charset = "utf-8"%>
<!--#include file="json_apiheader.asp"-->
<!--#include file="../inc/common.asp"-->
<!--#include file="../inc/Func.asp"-->
<!--#include file="../inc/JSON_2.0.4.asp"-->
<%
Dim jsa
status = "0000"
status_desc=""

select case request("act")
	case "get1"
		sqlstr = " select catego,convert(varchar,postdate,111) postdate,title,ison,attachfile1,attachname1 " & vbcrlf 
    sqlstr = sqlstr & " ,convert(varchar,createdate,111)+' '+convert(varchar,createdate,108) createdate " & vbcrlf 
    sqlstr = sqlstr & " from accountbbs " & vbcrlf 
    sqlstr = sqlstr & " where ison='Y' " & vbcrlf 
    sqlstr = sqlstr & " and catego='預算' " & vbcrlf 
    sqlstr = sqlstr & " order by postdate desc " & vbcrlf 
    set rs = server.createobject("adodb.recordset")
    rs.CursorLocation = 3
    rs.open sqlstr,objconn,3,1

    Set jsa = jsObject()
    jsa("status") = status
    jsa("status_desc") = status_desc
    set jsa("data") = jsArray()    
    
    do until rs.EOF
      set jsa("data")(null) = jsObject()
          
      for each x in rs.Fields
        jsa("data")(null)(x.name) = trim(x.value&"")
      next
      rs.movenext
    loop
    rs.close
    jsa.Flush
  case "get2"
    sqlstr = " select catego,convert(varchar,postdate,111) postdate,title,ison,attachfile1,attachname1 " & vbcrlf 
    sqlstr = sqlstr & " ,convert(varchar,createdate,111)+' '+convert(varchar,createdate,108) createdate " & vbcrlf 
    sqlstr = sqlstr & " from accountbbs " & vbcrlf 
    sqlstr = sqlstr & " where ison='Y' " & vbcrlf 
    sqlstr = sqlstr & " and catego='決算' " & vbcrlf 
    sqlstr = sqlstr & " order by postdate desc " & vbcrlf 
    set rs = server.createobject("adodb.recordset")
    rs.CursorLocation = 3
    rs.open sqlstr,objconn,3,1

    Set jsa = jsObject()
    jsa("status") = status
    jsa("status_desc") = status_desc
    set jsa("data") = jsArray()    
    
    do until rs.EOF
      set jsa("data")(null) = jsObject()
          
      for each x in rs.Fields
        jsa("data")(null)(x.name) = trim(x.value&"")
      next
      rs.movenext
    loop
    rs.close
    jsa.Flush
  case "get3"
    sqlstr = " select catego,convert(varchar,postdate,111) postdate,title,ison,attachfile1,attachname1 " & vbcrlf 
    sqlstr = sqlstr & " ,convert(varchar,createdate,111)+' '+convert(varchar,createdate,108) createdate " & vbcrlf 
    sqlstr = sqlstr & " from accountbbs " & vbcrlf 
    sqlstr = sqlstr & " where ison='Y' " & vbcrlf 
    sqlstr = sqlstr & " and catego='會計月報' " & vbcrlf 
    sqlstr = sqlstr & " order by postdate desc " & vbcrlf 
    set rs = server.createobject("adodb.recordset")
    rs.CursorLocation = 3
    rs.open sqlstr,objconn,3,1

    Set jsa = jsObject()
    jsa("status") = status
    jsa("status_desc") = status_desc
    set jsa("data") = jsArray()    
    
    do until rs.EOF
      set jsa("data")(null) = jsObject()
          
      for each x in rs.Fields
        jsa("data")(null)(x.name) = trim(x.value&"")
      next
      rs.movenext
    loop
    rs.close
    jsa.Flush    

end select


response.end%>