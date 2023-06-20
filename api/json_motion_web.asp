<%Response.CodePage = 65001%>
<%Response.Charset = "utf-8"%>
<%Response.ContentType= "application/json"%>
<!--#include file="json_apiheader.asp"-->
<!--#include file="../inc/common.asp"-->
<!--#include file="../inc/Func.asp"-->
<!--#include file="../inc/JSON_2.0.4.asp"-->
<%
Dim jsa
status = "0000"

select case request("act")
  case "getexpkd"

    sqlstr = "SELECT sm1_expkd FROM prtms_project WHERE sys_opsts<>'D' and sm1_publish=1 "
    sqlstr = sqlstr & " group by sm1_expkd " & vbcrlf
    sqlstr = sqlstr & " order by 1 desc " & vbcrlf
    set rs = objconn.execute(sqlstr)
    if not rs.eof then
      rtn_array = rs.getrows()
    end if
    rs.close
    set jsa = jsObject()
    jsa("status") = status
    jsa("status_desc") = status_desc
    set jsa("data") = jsArray()
    if isArray(rtn_array) then
      for rows=0 to ubound(rtn_array,2)
        set jsa("data")(null) = jsObject()
        jsa("data")(null)("value") = trim(rtn_array(0,rows)&"")
        jsa("data")(null)("label") = trim(rtn_array(0,rows)&"")
      next
    end if
    jsa.Flush

  case "getseqkd"

    sqlstr = "SELECT sm1_seqkd FROM prtms_project WHERE sys_opsts<>'D' and sm1_publish=1 "
    sqlstr = sqlstr & " group by sm1_seqkd " & vbcrlf
    sqlstr = sqlstr & " order by 1 desc " & vbcrlf
    set rs = objconn.execute(sqlstr)
    if not rs.eof then
      rtn_array = rs.getrows()
    end if
    rs.close
    set jsa = jsObject()
    jsa("status") = status
    jsa("status_desc") = status_desc
    set jsa("data") = jsArray()
    if isArray(rtn_array) then
      for rows=0 to ubound(rtn_array,2)
        set jsa("data")(null) = jsObject()
        jsa("data")(null)("value") = trim(rtn_array(0,rows)&"")

        label = ""
        select case trim(rtn_array(0,rows)&"")
          case "1"
            label = "定期會"
          case "2"
            label = "臨時會"
          case "3"
            label = "審查覆議案臨時會"  
        end select
        jsa("data")(null)("label") = label
      next
    end if
    jsa.Flush

case "getseqno"

    set sql_cmd = Server.CreateObject("ADODB.Command") 
    sql_cmd.ActiveConnection = Objconn

    sqlstr = " SELECT sm1_seqno,sm1_title " & vbcrlf 
    sqlstr = sqlstr & " FROM prtms_project  " & vbcrlf 
    sqlstr = sqlstr & " WHERE sys_opsts<>'D' and sm1_publish=1 " & vbcrlf 
    sqlstr = sqlstr & " and (? = '' or sm1_expkd= ?) " & vbcrlf 
    sqlstr = sqlstr & " and (? = '' or sm1_seqkd= ?) " & vbcrlf 
    sqlstr = sqlstr & " GROUP BY sm1_seqno,sm1_title " & vbcrlf 
    sqlstr = sqlstr & " ORDER BY 1 desc " & vbcrlf 

    sql_cmd.CommandText = sqlstr

    expkd = ""
    if request("expkd")>"" then
      expkd = trim(request("expkd"))
    end if

    seqkd = ""
    if request("seqkd")>"" then
      seqkd = trim(request("seqkd"))
    end if

    'ADO.CreateParameter(name,type,direction,size,value)
    sql_cmd.Parameters.Append sql_cmd.CreateParameter("sm1_expkd1",202,1,20,expkd)
    sql_cmd.Parameters.Append sql_cmd.CreateParameter("sm1_expkd2",202,1,20,expkd)
    sql_cmd.Parameters.Append sql_cmd.CreateParameter("sm1_seqkd1",202,1,20,seqkd)
    sql_cmd.Parameters.Append sql_cmd.CreateParameter("sm1_seqkd2",202,1,20,seqkd)
    set rs = sql_cmd.Execute

    if not rs.eof then
      rtn_array = rs.getrows()
    end if
    rs.close
    set jsa = jsObject()
    jsa("status") = status
    jsa("status_desc") = status_desc
    set jsa("data") = jsArray()
    if isArray(rtn_array) then
      for rows=0 to ubound(rtn_array,2)
        set jsa("data")(null) = jsObject()
        jsa("data")(null)("value") = trim(rtn_array(0,rows)&"")
        jsa("data")(null)("label") = trim(rtn_array(1,rows)&"")
      next
    end if
    jsa.Flush

case "getlist"

    set sql_cmd = Server.CreateObject("ADODB.Command") 
    sql_cmd.ActiveConnection = Objconn

    sqlstr = " SELECT sm1_expkd,sm1_seqkd,sm1_seqno,sm1_title " & vbcrlf 
    sqlstr = sqlstr & " ,sm1_3c   " & vbcrlf 
    sqlstr = sqlstr & " ,sm1_7c  " & vbcrlf 
    sqlstr = sqlstr & " ,sm1_5c  " & vbcrlf 
    sqlstr = sqlstr & " ,sm1_1c  " & vbcrlf 
    sqlstr = sqlstr & " ,sm1_2c  " & vbcrlf 
    sqlstr = sqlstr & " FROM prtms_project  " & vbcrlf 
    sqlstr = sqlstr & " WHERE sys_opsts<>'D' and sm1_publish=1 " & vbcrlf 
    if request("expkd")>"" then
      sqlstr = sqlstr & " and sm1_expkd= ? " & vbcrlf   
    end if
    if request("seqkd")>"" then
      sqlstr = sqlstr & " and sm1_seqkd= ? " & vbcrlf   
    end if
    if request("seqno")>"" then
      sqlstr = sqlstr & " and sm1_title= ? " & vbcrlf  
    end if

    if request("searchstr")>"" then
      sqlstr = sqlstr & " and (sm1_3c like ? or sm1_5c like ? or sm1_7c like ?) " & vbcrlf 
    end if

    if request("sm1_1c")>"" then
      sqlstr = sqlstr & " and sm1_1C like ? " & vbcrlf  
    end if

    if request("sm1_2c")>"" then
      sqlstr = sqlstr & " and sm1_2C like ? " & vbcrlf  
    end if


    sqlstr = sqlstr & " ORDER BY sm1_expkd desc, sm1_seqkd desc , sm1_seqno desc  " & vbcrlf 

    sql_cmd.CommandText = sqlstr

    expkd = ""
    if request("expkd")>"" then
      expkd = request("expkd")
      sql_cmd.Parameters.Append sql_cmd.CreateParameter("sm1_expkd",202,1,20,expkd)
    end if

    seqkd = ""
    if request("seqkd")>"" then
      seqkd = request("seqkd")
      sql_cmd.Parameters.Append sql_cmd.CreateParameter("sm1_seqkd",202,1,20,seqkd)
    end if

    seqno = ""
    if request("seqno")>"" then
      seqno = trim(request("seqno"))
      sql_cmd.Parameters.Append sql_cmd.CreateParameter("sm1_title",202,1,100,seqno)
    end if

    searchstr = ""
    if request("searchstr")>"" then
      searchstr = "%"&trim(request("searchstr"))&"%"
      sql_cmd.Parameters.Append sql_cmd.CreateParameter("search1",202,1,100,searchstr)
      sql_cmd.Parameters.Append sql_cmd.CreateParameter("search2",202,1,100,searchstr)
      sql_cmd.Parameters.Append sql_cmd.CreateParameter("search3",202,1,100,searchstr)      
    end if


    sm1_1C = ""
    if request("sm1_1c")>"" then
      sm1_1C = "%"&trim(request("sm1_1c"))&"%"
      sql_cmd.Parameters.Append sql_cmd.CreateParameter("sm1_1c",202,1,50,sm1_1C)
    end if

    sm1_2C = ""
    if request("sm1_2c")>"" then
      sm1_2C = "%"&trim(request("sm1_2c"))&"%"
      sql_cmd.Parameters.Append sql_cmd.CreateParameter("sm1_2c",202,1,50,sm1_2C)
    end if

    'ADO.CreateParameter(name,type,direction,size,value)

    set rs = sql_cmd.Execute

    if not rs.eof then
      rtn_array = rs.getrows()
    end if
    rs.close
    set jsa = jsObject()
    jsa("status") = status
    jsa("status_desc") = status_desc
    set jsa("data") = jsArray()
    if isArray(rtn_array) then
      for rows=0 to ubound(rtn_array,2)
        set jsa("data")(null) = jsObject()
        jsa("data")(null)("sm1_expkd") = trim(rtn_array(0,rows)&"")
        jsa("data")(null)("sm1_seqkd") = trim(rtn_array(1,rows)&"")
        jsa("data")(null)("sm1_seqno") = trim(rtn_array(2,rows)&"")
        jsa("data")(null)("sm1_title") = trim(rtn_array(3,rows)&"")
        jsa("data")(null)("sm1_3c") = trim(rtn_array(4,rows)&"")
        jsa("data")(null)("sm1_7c") = trim(rtn_array(5,rows)&"")
        jsa("data")(null)("sm1_5c") = trim(rtn_array(6,rows)&"")
        jsa("data")(null)("sm1_1c") = trim(rtn_array(7,rows)&"")
        jsa("data")(null)("sm1_2c") = trim(rtn_array(8,rows)&"")
      next
    end if
    jsa.Flush    

end select


response.end%>