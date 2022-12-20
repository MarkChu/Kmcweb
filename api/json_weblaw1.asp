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
	case "getdet"

		sqlstr = " select uniqid as detid,lawid,chcatego,chtitle,chcontent,sortid " & vbcrlf 
    sqlstr = sqlstr & " from weblawdet " & vbcrlf 
    sqlstr = sqlstr & " where lawid='"&request("lawid")&"' " & vbcrlf 
    sqlstr = sqlstr & " order by sortid,chcatego, chtitle " & vbcrlf 
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