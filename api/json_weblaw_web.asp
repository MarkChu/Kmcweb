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
  case "getdetail"

    set sql_cmd = Server.CreateObject("ADODB.Command") 
    sql_cmd.ActiveConnection = Objconn

    sqlstr = " select lawid,lawcatego,lawtitle,lawcontent,[url] " & vbcrlf 
    sqlstr = sqlstr & " from weblaw " & vbcrlf 
    sqlstr = sqlstr & " where ison='Y' and lawid = ? " & vbcrlf 

    sql_cmd.CommandText = sqlstr

    lawid = ""
    if request("lawid")>"" then
      lawid = trim(request("lawid"))
    end if

    sql_cmd.Parameters.Append sql_cmd.CreateParameter("lawid",202,1,50,lawid)
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
        lawid = trim(rtn_array(0,rows)&"")
        jsa("data")(null)("lawid") = lawid
        jsa("data")(null)("lawcatego") = trim(rtn_array(1,rows)&"")
        select case trim(rtn_array(1,rows)&"")
          case "1"
            lawcategoname = "地方議政法規"
          case "2"
            lawcategoname = "相關法令解釋"
          case "3"
            lawcategoname = "財政預算法規"
          case "4"
            lawcategoname = "其他法規"
        end select
        jsa("data")(null)("lawcategoname") = lawcategoname
        jsa("data")(null)("lawtitle") = trim(rtn_array(2,rows)&"")
        jsa("data")(null)("lawcontent") = trim(rtn_array(3,rows)&"")
        jsa("data")(null)("url") = trim(rtn_array(4,rows)&"")


        set jsa("data")(null)("chcatego") = jsArray()
        set ch_array = nothing
        sqlstr = " select chcatego,min(sortid) sortid " & vbcrlf 
        sqlstr = sqlstr & " from weblawdet " & vbcrlf 
        sqlstr = sqlstr & " where lawid='"&lawid&"' " & vbcrlf 
        sqlstr = sqlstr & " group by chcatego " & vbcrlf 
        sqlstr = sqlstr & " order by 2 " & vbcrlf 
        set rs = objconn.execute(sqlstr)
        if not rs.eof then
          ch_array = rs.getrows()
        end if
        rs.close

        if isArray(ch_array) then

          for chn=0 to ubound(ch_array,2)
            set jsa("data")(null)("chcatego")(null) = jsObject()
            jsa("data")(null)("chcatego")(null)("id") = chn + 1
            jsa("data")(null)("chcatego")(null)("chname") = trim(ch_array(0,chn))

            set jsa("data")(null)("chcatego")(null)("chdata") = jsArray()

            sqlstr = " select chtitle,chcontent " & vbcrlf 
            sqlstr = sqlstr & " from weblawdet " & vbcrlf 
            sqlstr = sqlstr & " where lawid="&lawid&" " & vbcrlf 
            sqlstr = sqlstr & " and isnull(chcatego,'')='"&trim(ch_array(0,chn))&"' " & vbcrlf 
            sqlstr = sqlstr & " order by sortid " & vbcrlf 
            set chdet_array = nothing
            set rs = objconn.execute(sqlstr)
            if not rs.eof then
              chdet_array = rs.getrows()
            end if
            rs.close            

            if isArray(chdet_array) then
              for chx = 0 to ubound(chdet_array,2)
                set jsa("data")(null)("chcatego")(null)("chdata")(null) = jsObject()
                jsa("data")(null)("chcatego")(null)("chdata")(null)("id") = chx + 1
                jsa("data")(null)("chcatego")(null)("chdata")(null)("chtitle") = trim(chdet_array(0,chx))
                jsa("data")(null)("chcatego")(null)("chdata")(null)("chcontent") = trim(chdet_array(1,chx))
              next
            end if

          next
        end if

      next
    end if
    jsa.Flush

  case "search"

    set sql_cmd = Server.CreateObject("ADODB.Command") 
    sql_cmd.ActiveConnection = Objconn

    sqlstr = " declare @str as nvarchar(50) " & vbcrlf 
    sqlstr = sqlstr & " set @str = ? " & vbcrlf 
    sqlstr = sqlstr & " select lawid,lawcatego,lawtitle,url " & vbcrlf 
    sqlstr = sqlstr & " ,case when lawcatego=1 then chcatego + ' - ' + chtitle " & vbcrlf 
    sqlstr = sqlstr & "     else LawTitle " & vbcrlf 
    sqlstr = sqlstr & " end as detcontent " & vbcrlf 
    sqlstr = sqlstr & " from  " & vbcrlf 
    sqlstr = sqlstr & " ( " & vbcrlf 
    sqlstr = sqlstr & "   select a.lawid,lawcatego,lawtitle,b.chcatego,a.url,a.sortid,b.chtitle,b.chcontent,a.LawContent " & vbcrlf 
    sqlstr = sqlstr & "   from weblaw a left outer join weblawdet b " & vbcrlf 
    sqlstr = sqlstr & "   on a.lawid = b.lawid " & vbcrlf 
    sqlstr = sqlstr & "   and a.IsOn='Y' " & vbcrlf 
    sqlstr = sqlstr & "   and a.lawcatego in (1,2) " & vbcrlf 
    sqlstr = sqlstr & "   where (a.LawTitle like @str " & vbcrlf 
    sqlstr = sqlstr & "     or a.LawContent like @str " & vbcrlf 
    sqlstr = sqlstr & "     or b.ChTitle like @str " & vbcrlf 
    sqlstr = sqlstr & "     or b.ChContent like @str " & vbcrlf 
    sqlstr = sqlstr & "   ) " & vbcrlf 
    sqlstr = sqlstr & "   group by a.lawid,lawcatego,lawtitle,b.chcatego,a.url,a.sortid,b.chtitle,b.chcontent,a.LawContent " & vbcrlf 
    sqlstr = sqlstr & " ) a " & vbcrlf 
    sqlstr = sqlstr & " order by a.lawcatego, a.sortid, a.ChCatego " & vbcrlf 


    sql_cmd.CommandText = sqlstr

    searchstr = ""
    if request("searchstr")>"" then
      searchstr = "%"&trim(request("searchstr"))&"%"
    else
      searchstr = "ZZZYYYDDD"
    end if

    'ADO.CreateParameter(name,type,direction,size,value)
    sql_cmd.Parameters.Append sql_cmd.CreateParameter("@str",202,1,40,searchstr)
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
        jsa("data")(null)("lawid") = trim(rtn_array(0,rows)&"")
        jsa("data")(null)("lawcatego") = trim(rtn_array(1,rows)&"")
        select case trim(rtn_array(1,rows)&"")
          case "1"
            lawcategoname = "地方議政法規"
          case "2"
            lawcategoname = "相關法令解釋"
          case "3"
            lawcategoname = "財政預算法規"
          case "4"
            lawcategoname = "其他法規"
        end select
        jsa("data")(null)("lawcategoname") = lawcategoname
        jsa("data")(null)("lawtitle") = trim(rtn_array(2,rows)&"")
        jsa("data")(null)("url") = trim(rtn_array(3,rows)&"")
        jsa("data")(null)("searchcontent") = trim(rtn_array(4,rows)&"")
      next
    end if
    jsa.Flush

case "getlist1","getlist2","getlist3","getlist4"

    set sql_cmd = Server.CreateObject("ADODB.Command") 
    sql_cmd.ActiveConnection = Objconn

    sqlstr = " select a.lawid,a.lawcatego,lawtitle,a.url,a.lawcontent " & vbcrlf 
    sqlstr = sqlstr & " from weblaw a " & vbcrlf 
    sqlstr = sqlstr & " where a.IsOn='Y' " & vbcrlf 
    select case request("act")
      case "getlist1"
        sqlstr = sqlstr & " and a.lawcatego=1 " & vbcrlf 
      case "getlist2"
        sqlstr = sqlstr & " and a.lawcatego=2 " & vbcrlf 
      case "getlist3"
        sqlstr = sqlstr & " and a.lawcatego=3 " & vbcrlf       
      case "getlist4"
        sqlstr = sqlstr & " and a.lawcatego=4 " & vbcrlf       
    end select
    sqlstr = sqlstr & " order by a.sortid " & vbcrlf 
    sql_cmd.CommandText = sqlstr
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
        jsa("data")(null)("lawid") = trim(rtn_array(0,rows)&"")
        jsa("data")(null)("lawcatego") = trim(rtn_array(1,rows)&"")
        select case trim(rtn_array(1,rows)&"")
          case "1"
            lawcategoname = "地方議政法規"
          case "2"
            lawcategoname = "相關法令解釋"
          case "3"
            lawcategoname = "財政預算法規"
          case "4"
            lawcategoname = "其他法規"
        end select
        jsa("data")(null)("lawcategoname") = lawcategoname
        jsa("data")(null)("lawtitle") = trim(rtn_array(2,rows)&"")
        jsa("data")(null)("url") = trim(rtn_array(3,rows)&"")
        jsa("data")(null)("lawcontent") = trim(rtn_array(4,rows)&"")
      next
    end if
    jsa.Flush

end select


response.end%>