<%Response.CodePage = 65001%>
<%Response.Charset = "utf-8"%>
<%Response.ContentType= "application/json"%>
<!--#include file="json_common.asp"-->
<!--#include file="../inc/common.asp"-->
<!--#include file="../inc/Func.asp"-->
<!--#include file="../inc/JSON_2.0.4.asp"-->
<%
Dim jsa
status = "0000"

select case request("act")
	case "getdtl"
		sm1_ln = request("sm1_ln")
       	querystr = querystr & " AND sm1_ln='"&sm1_ln&"' "

	    sqlstr = "SELECT * FROM prtms_project WHERE sys_opsts<>'D' "
	    sqlstr = sqlstr & querystr
	    sqlstr = sqlstr & " ORDER BY sm1_chkkd DESC" & vbcrlf
      	Set jsa = jsObject()
		jsa("status") = status
		jsa("status_desc") = status_desc
		set jsa("data") = jsArray()

      	set rs = server.createobject("adodb.recordset")
      	rs.CursorLocation = 3
      	rs.open sqlstr,objconn,3,1
      	do until rs.EOF
      		set jsa("data")(null) = jsObject()
      		
			for each x in rs.Fields
				jsa("data")(null)(x.name) = trim(x.value&"")
  			next
      		jsa("data")(null)("sm1_1C") = replace(trim(rs("sm1_1C")&""),vbcrlf,"、")
      		jsa("data")(null)("sm1_2C") = replace(trim(rs("sm1_2C")&""),vbcrlf,"、")
      		jsa("data")(null)("sm1_3C") = replace(trim(rs("sm1_3C")&""),vbcrlf,"<br>")
      		jsa("data")(null)("sm1_4C") = replace(trim(rs("sm1_4C")&""),vbcrlf,"<br>")
      		jsa("data")(null)("sm1_5C") = replace(trim(rs("sm1_5C")&""),vbcrlf,"<br>")
      		jsa("data")(null)("sm1_6C") = replace(trim(rs("sm1_6C")&""),vbcrlf,"<br>")
      		jsa("data")(null)("sm1_7C") = replace(trim(rs("sm1_7C")&""),vbcrlf,"<br>")
  			rs.movenext
  		loop
      	rs.close   	
      	jsa.Flush


	case "getlist"
		sm1_expkd = request("sm1_expkd")
		sm1_seqkd = request("sm1_seqkd")
		sm1_title = request("sm1_title")		
		sm1_seqno = request("sm1_seqno")
		sm1_sugkd = request("sm1_sugkd")
		sm1_modkd = request("sm1_modkd")
		sm1_chkkd = request("sm1_chkkd")
		search_content = request("search_content")
		sm1_1C = request("sm1_1C")
		sm1_2C = request("sm1_2C")

		if sm1_expkd>"" then
        	querystr = querystr & " AND sm1_expkd='"&sm1_expkd&"' "
    	end if
		if sm1_seqkd>"" then
        	querystr = querystr & " AND sm1_seqkd='"&sm1_seqkd&"' "
    	end if
		if sm1_seqno>"" then
        	querystr = querystr & " AND sm1_seqno='"&sm1_seqno&"' "
    	end if    
		if sm1_sugkd>"" then
        	querystr = querystr & " AND sm1_sugkd='"&sm1_sugkd&"' "
    	end if
		if sm1_modkd>"" then
        	querystr = querystr & " AND sm1_modkd='"&sm1_modkd&"' "
    	end if
		if sm1_chkkd>"" then
        	querystr = querystr & " AND sm1_chkkd='"&sm1_chkkd&"' "
    	end if
		if search_content>"" then
        	querystr = querystr & " AND (sm1_3C like '%"&search_content&"%' or sm1_4C like '%"&search_content&"%' or sm1_5C like '%"&search_content&"%' or sm1_6C like '%"&search_content&"%' or sm1_7C like '%"&search_content&"%' or sm1_8C like '%"&search_content&"%') "
    	end if
		if sm1_1C>"" then
        	querystr = querystr & " AND sm1_1C like '"&sm1_1C&"' "
    	end if
    	if sm1_2C>"" then
        	querystr = querystr & " AND sm1_2C like '"&sm1_2C&"' "
    	end if
		if sm1_title>"" then
        	querystr = querystr & " AND sm1_title like '"&sm1_title&"' "
    	end if
    	    	    	    
	    sqlstr = "SELECT * FROM prtms_project WHERE sys_opsts<>'D' "
	    sqlstr = sqlstr & querystr
	    sqlstr = sqlstr & " ORDER BY sm1_chkkd DESC" & vbcrlf


      	Set jsa = jsObject()
		jsa("status") = status
		jsa("status_desc") = status_desc
		set jsa("data") = jsArray()

      	set rs = server.createobject("adodb.recordset")
      	rs.CursorLocation = 3
      	rs.open sqlstr,objconn,3,1
      	do until rs.EOF
      		set jsa("data")(null) = jsObject()
      		set jsa("data")(null)("options") = jsObject()
      		jsa("data")(null)("options")("classes")="ft-body-row"
      		set jsa("data")(null)("value") = jsObject()      		

      		set jsa("data")(null)("value")("uniqid")= jsObject()
      		set jsa("data")(null)("value")("uniqid")("options")=jsObject()
      		jsa("data")(null)("value")("uniqid")("options")("visible")="true"
      		jsa("data")(null)("value")("uniqid")("value")=rs("sm1_ln")

  			for each x in rs.Fields
  				select case x.name
  					case "sm1_seqno","sm1_sugno"
			      		set jsa("data")(null)("value")(x.name)= jsObject()
			      		set jsa("data")(null)("value")(x.name)("options")=jsObject()
			      		v = 0
			      		if isnumeric(trim(x.value&"")) then
			      			v = cdbl(trim(x.value&""))
			      		end if
			      		jsa("data")(null)("value")(x.name)("options")("sortValue")=v
			      		jsa("data")(null)("value")(x.name)("value") = trim(x.value&"")
  					case else
  						jsa("data")(null)("value")(x.name) = trim(x.value&"")
  				end select
  				
  			next
  			rs.movenext
  		loop
      	rs.close   	
      	jsa.Flush

	case "getseq"
		sm1_expkd = request("sm1_expkd")
		sm1_seqkd = request("sm1_seqkd")
		sm1_title = request("sm1_title")
		if sm1_expkd>"" then
        	querystr = querystr & " AND sm1_expkd='"&sm1_expkd&"' "
    	end if
		if sm1_seqkd>"" then
        	querystr = querystr & " AND sm1_seqkd='"&sm1_seqkd&"' "
    	end if
    
	    sqlstr = "SELECT sm1_title,sm1_seqno FROM prtms_project WHERE sys_opsts<>'D' "
	    sqlstr = sqlstr & querystr
	    sqlstr = sqlstr & " GROUP BY sm1_title,sm1_seqno " & vbcrlf
	    sqlstr = sqlstr & " ORDER BY 2 DESC " & vbcrlf
      	set rs = objconn.execute(sqlstr)
      	if not rs.eof then
      		rtn_array = rs.getrows()
      	end if
      	rs.close

      	Set jsa = jsObject()
		jsa("status") = status
		jsa("status_desc") = status_desc

		set jsa("data") = jsArray()
		if isArray(rtn_array) then
			for rows=0 to ubound(rtn_array,2)
				set jsa("data")(null) = jsObject()
				jsa("data")(null)("sm1_title") = trim(rtn_array(0,rows)&"")
				jsa("data")(null)("sm1_seqno") = rtn_array(1,rows)
			next
		end if
		jsa.Flush

  case "getseqD"

    sm1_expkd = request("sm1_expkd")
    sm1_seqkd = request("sm1_seqkd")

    if sm1_expkd>"" then
          querystr = querystr & " AND GN="&sm1_expkd&" "
      end if
    if sm1_seqkd>"" then
          querystr = querystr & " AND HB="&sm1_seqkd&" "
      end if
    
      sqlstr = "SELECT HN as sm1_seqno FROM prtms_project_D1114 WHERE 1=1 "
      sqlstr = sqlstr & querystr
      sqlstr = sqlstr & " GROUP BY HN " & vbcrlf
      sqlstr = sqlstr & " ORDER BY 1 DESC " & vbcrlf
        set rs = objconn.execute(sqlstr)
        if not rs.eof then
          rtn_array = rs.getrows()
        end if
        rs.close

        Set jsa = jsObject()
    jsa("status") = status
    jsa("status_desc") = status_desc

    set jsa("data") = jsArray()
    if isArray(rtn_array) then
      for rows=0 to ubound(rtn_array,2)
        set jsa("data")(null) = jsObject()
        jsa("data")(null)("sm1_title") = trim(rtn_array(0,rows)&"")
        jsa("data")(null)("sm1_seqno") = rtn_array(0,rows)
      next
    end if
    jsa.Flush

  case "getlistD"
    sm1_expkd = request("sm1_expkd")
    sm1_seqkd = request("sm1_seqkd")
    sm1_seqno = request("sm1_seqno")
    sm1_modkd = request("sm1_modkd")
    sm1_chkkd = request("sm1_chkkd")
    search_content = request("search_content")


    if sm1_expkd>"" then
          querystr = querystr & " AND GN="&sm1_expkd&" "
    end if
    if sm1_seqkd>"" then
          querystr = querystr & " AND HB="&sm1_seqkd&" "
    end if
    if sm1_seqno>"" then
          querystr = querystr & " AND HN="&sm1_seqno&" "
    end if    
    if sm1_modkd>"" then
          querystr = querystr & " AND T="&sm1_modkd&" "
    end if
    if sm1_chkkd>"" then
          querystr = querystr & " AND KIND='"&sm1_chkkd&"' "
    end if
    if search_content>"" then
          querystr = querystr & " AND (EW like '%"&search_content&"%' or EH like '%"&search_content&"%' or EO like '%"&search_content&"%') "
    end if

                      
      sqlstr = "SELECT GN as sm1_expkd " & vbcrlf
      sqlstr = sqlstr & " ,HB as sm1_seqkd " & vbcrlf
      sqlstr = sqlstr & " ,HN as sm1_seqno " & vbcrlf
      sqlstr = sqlstr & " ,T as sm1_modkd " & vbcrlf
      sqlstr = sqlstr & " ,KIND as sm1_chkkd " & vbcrlf
      sqlstr = sqlstr & " ,ID as sm1_ln " & vbcrlf
      sqlstr = sqlstr & " ,SN as sm1_id " & vbcrlf
      sqlstr = sqlstr & " FROM prtms_project_D1114 WHERE 1=1 " & vbcrlf
      sqlstr = sqlstr & querystr
      sqlstr = sqlstr & " ORDER BY KIND DESC " & vbcrlf


    Set jsa = jsObject()
    jsa("status") = status
    jsa("status_desc") = status_desc
    set jsa("data") = jsArray()

        set rs = server.createobject("adodb.recordset")
        rs.CursorLocation = 3
        rs.open sqlstr,objconn,3,1
        do until rs.EOF
          set jsa("data")(null) = jsObject()
          set jsa("data")(null)("options") = jsObject()
          jsa("data")(null)("options")("classes")="ft-body-row"
          set jsa("data")(null)("value") = jsObject()         

          set jsa("data")(null)("value")("uniqid")= jsObject()
          set jsa("data")(null)("value")("uniqid")("options")=jsObject()
          jsa("data")(null)("value")("uniqid")("options")("visible")="true"
          jsa("data")(null)("value")("uniqid")("value")=rs("sm1_ln")

        for each x in rs.Fields
          select case x.name
            case "sm1_seqno","sm1_sugno"
                set jsa("data")(null)("value")(x.name)= jsObject()
                set jsa("data")(null)("value")(x.name)("options")=jsObject()
                v = 0
                if isnumeric(trim(x.value&"")) then
                  v = cdbl(trim(x.value&""))
                end if
                jsa("data")(null)("value")(x.name)("options")("sortValue")=v
                jsa("data")(null)("value")(x.name)("value") = trim(x.value&"")
            case else
              jsa("data")(null)("value")(x.name) = trim(x.value&"")
          end select
          
        next
        rs.movenext
      loop
        rs.close    
        jsa.Flush

  case "getdtlD"

    sm1_ln = request("sm1_ln")
    querystr = querystr & " AND ID='"&sm1_ln&"' "

    sqlstr = "SELECT GN as sm1_expkd " & vbcrlf
    sqlstr = sqlstr & " ,HB as sm1_seqkd " & vbcrlf
    sqlstr = sqlstr & " ,HN as sm1_seqno " & vbcrlf
    sqlstr = sqlstr & " ,T as sm1_modkd " & vbcrlf
    sqlstr = sqlstr & " ,KIND as sm1_chkkd " & vbcrlf
    sqlstr = sqlstr & " ,ID as sm1_ln " & vbcrlf
    sqlstr = sqlstr & " ,EW as sm1_3C " & vbcrlf
    sqlstr = sqlstr & " ,EH as sm1_1C " & vbcrlf
    sqlstr = sqlstr & " ,EO as sm1_8C " & vbcrlf
    sqlstr = sqlstr & " ,SN as sm1_id " & vbcrlf
    sqlstr = sqlstr & " FROM prtms_project_D1114 WHERE 1=1 " & vbcrlf
    sqlstr = sqlstr & querystr



    sqlstr = sqlstr & querystr
    Set jsa = jsObject()
    jsa("status") = status
    jsa("status_desc") = status_desc
    set jsa("data") = jsArray()

    set rs = server.createobject("adodb.recordset")
    rs.CursorLocation = 3
    rs.open sqlstr,objconn,3,1
    do until rs.EOF
      set jsa("data")(null) = jsObject()

      for each x in rs.Fields
        jsa("data")(null)(x.name) = trim(x.value&"")
      next
        jsa("data")(null)("sm1_1C") = replace(trim(rs("sm1_1C")&""),vbcrlf,"、")
        jsa("data")(null)("sm1_2C") = ""
        jsa("data")(null)("sm1_3C") = replace(trim(rs("sm1_3C")&""),vbcrlf,"<br>")
        jsa("data")(null)("sm1_4C") = ""
        jsa("data")(null)("sm1_5C") = ""
        jsa("data")(null)("sm1_6C") = ""
        jsa("data")(null)("sm1_7C") = ""
        jsa("data")(null)("sm1_8C") = replace(trim(rs("sm1_8C")&""),vbcrlf,"<br>")
      rs.movenext
    loop
    rs.close    
    jsa.Flush

  case "getlistK"
    HN = request("HN")
 
    if HN>"" then
      querystr = querystr & " AND HN='"&HN&"' "
    end if

    sqlstr = "SELECT hn as sm1_book,kind as sm1_page,img as sm1_img  " & vbcrlf
    sqlstr = sqlstr & " FROM prtms_project_D0110 WHERE 1=1 " & vbcrlf
    sqlstr = sqlstr & querystr
    sqlstr = sqlstr & " ORDER BY kind " & vbcrlf


    Set jsa = jsObject()
    jsa("status") = status
    jsa("status_desc") = status_desc
    set jsa("data") = jsArray()

    set rs = server.createobject("adodb.recordset")
    rs.CursorLocation = 3
    rs.open sqlstr,objconn,3,1
    do until rs.EOF

      set jsa("data")(null) = jsObject()
      set jsa("data")(null)("options") = jsObject()
      jsa("data")(null)("options")("classes")="ft-body-row"
      set jsa("data")(null)("value") = jsObject()         

      for each x in rs.Fields
        jsa("data")(null)("value")(x.name) = trim(x.value&"")
      next
      rs.movenext
    loop
    rs.close    
    jsa.Flush


end select


response.end%>