<%Response.CodePage = 65001%>
<%Response.Charset = "utf-8"%>
<%Response.ContentType= "application/json"%>
<!--#include file="json_common.asp"-->
<!--#include file="../inc/common.asp"-->
<!--#include file="../inc/Func.asp"-->
<!--#include file="../inc/JSON_2.0.4.asp"-->
<%
status = "0000"
userid = request("userid")
topeople = request("topeople")
topeople_txt = request("topeople_txt")
tomedia = request("tomedia")
tomedia_txt = request("tomedia_txt")
sms_body = request("sms_body")

'全部媒體
allmedia = request("allmedia")

send_y = request("send_y")
send_m = request("send_m")
send_d = request("send_d")
send_h = request("send_h")
send_n = request("send_n")
send_s = request("send_s")

send_time = send_y&"/"&send_m&"/"&send_d&" "&send_h&":"&send_n&":"&send_s

if not isdate(send_time) then
	status = "9999"
	status_desc = "簡訊發送日期錯誤!!"
end if

if topeople>"" then
	people_ary = split(topeople,",")
	people_mpt_ary = split(topeople_txt,",")
end if

if tomedia>"" then
	media_ary = split(tomedia,",")
	media_mpt_ary = split(tomedia_txt,",")
end if


if status="0000" then

	sqlstr = "select newid() "
	set rs = objconn.execute(sqlstr)
	if not rs.eof then
		group_id = trim(rs(0))
	end if
	rs.close

	Objconn.BeginTrans


	if isArray(people_ary) then
		for r=0 to ubound(people_ary)
			if trim(people_ary(r)&"")>"" then


				item_str = trim(people_ary(r))
				item_ary = split(item_str,"@")
				isdept = false
				if trim(item_ary(1)&"")="" then
					isdept = true
				end if

				if isdept then
					sqlstr = "select account_id,[Name],[cell] from vw_Members_Login where cell>'' and len(cell)=10	and dept_id='"&trim(item_ary(0))&"'"
				else 
					sqlstr = "select account_id,[Name],[cell] from vw_Members_Login where cell>'' and len(cell)=10	and account_id='"&trim(item_ary(0))&"'"
				end if
				set rs = objconn.execute(sqlstr)
				if not rs.eof then
					memb_ary = rs.getrows()
				else 
					if isArray(memb_ary) then
						set memb_ary = nothing
					end if
				end if
				rs.close
				if IsArray(memb_ary) then
					for rows=0 to ubound(memb_ary,2)
						account_id = trim(memb_ary(0,rows)&"")
						account_name = trim(memb_ary(1,rows)&"")
						cell_no = trim(memb_ary(2,rows)&"")

						sqlstr = "select max(unique_id) from sms_result "
						set rs = objconn.execute(sqlstr)
						if not rs.eof then
							det_unique_id = cdbl("0"&rs(0)) + 1
						else
							det_unique_id =1
						end if
						rs.close

						isInsert = false

						sqlstr = "select count(*) from sms_result where group_id='"&group_id&"' and people='"&account_id&"'"
						set rs = ObjConn.execute(sqlstr)
						if cdbl(rs(0))=0 then
							isInsert = true
						end if
						rs.close

						if isInsert then


							sqlstr = "select max(unique_id) from sms_result"
							set rs = objconn.execute(sqlstr)
							if not rs.eof then
								unique_id = cdbl("0"&rs(0)) + 1
							else
								unique_id =1
							end if
							rs.close

							sqlstr = "insert into SMS_Result (unique_id " & vbcrlf 
							sqlstr = sqlstr & ",group_id " & vbcrlf 
							sqlstr = sqlstr & ",cell_no " & vbcrlf 
							sqlstr = sqlstr & ",[description] " & vbcrlf 
							sqlstr = sqlstr & ",send_time " & vbcrlf 
							sqlstr = sqlstr & ",sent_result " & vbcrlf 
							sqlstr = sqlstr & ",people_name " & vbcrlf 
							sqlstr = sqlstr & ",people " & vbcrlf 
							sqlstr = sqlstr & ",postdate " & vbcrlf 
							sqlstr = sqlstr & ",account_name " & vbcrlf 
							sqlstr = sqlstr & ",account_id " & vbcrlf 
							sqlstr = sqlstr & ",system_id " & vbcrlf 
							sqlstr = sqlstr & ",delta_day " & vbcrlf 
							sqlstr = sqlstr & ",delta_hour " & vbcrlf 
							sqlstr = sqlstr & ",delta_minute " & vbcrlf 
							sqlstr = sqlstr & ",visible_flag " & vbcrlf 
							sqlstr = sqlstr & ") values ("&det_unique_id&" " & vbcrlf 
							sqlstr = sqlstr & ","&FnSQL(group_id,0)&" " & vbcrlf 
							sqlstr = sqlstr & ","&FnSQL(cell_no,0)&" " & vbcrlf 
							sqlstr = sqlstr & ","&FnSQL(sms_body,0)&" " & vbcrlf 
							sqlstr = sqlstr & ","&FnSQL(send_time,2)&" " & vbcrlf 
							sqlstr = sqlstr & ",0 " & vbcrlf 
							sqlstr = sqlstr & ","&FnSQL(account_name,0)&" " & vbcrlf 
							sqlstr = sqlstr & ","&FnSQL(account_id,0)&" " & vbcrlf 
							sqlstr = sqlstr & ",getdate() " & vbcrlf 
							sqlstr = sqlstr & ","&FnSQL(session("username"),0)&" " & vbcrlf 
							sqlstr = sqlstr & ","&FnSQL(session("userid"),0)&" " & vbcrlf 
							sqlstr = sqlstr & ",1 " & vbcrlf 						
							sqlstr = sqlstr & ",0 " & vbcrlf 
							sqlstr = sqlstr & ",0 " & vbcrlf 
							sqlstr = sqlstr & ",0 " & vbcrlf 
							sqlstr = sqlstr & ",1 " & vbcrlf 
							sqlstr = sqlstr & ") " & vbcrlf 
							objconn.execute(sqlstr)


						end if

					next
				end if


			end if
		next
	end if


	if allmedia>"" then
		'20210405 全部媒體

		sqlstr = "select unique_id,[Name],[cell] from members1 where cell>'' and len(cell)=10"
		set rs = objconn.execute(sqlstr)
		if not rs.eof then
			memb_ary = rs.getrows()
		else 
			if isArray(memb_ary) then
				set memb_ary = nothing
			end if
		end if
		rs.close
		if IsArray(memb_ary) then
			for rows=0 to ubound(memb_ary,2)
				account_id = trim(memb_ary(0,rows)&"")
				account_name = trim(memb_ary(1,rows)&"")
				cell_no = trim(memb_ary(2,rows)&"")

				sqlstr = "select max(unique_id) from sms_result "
				set rs = objconn.execute(sqlstr)
				if not rs.eof then
					det_unique_id = cdbl("0"&rs(0)) + 1
				else
					det_unique_id =1
				end if
				rs.close

				isInsert = false

				sqlstr = "select count(*) from sms_result where group_id='"&group_id&"' and people='"&account_id&"'"
				set rs = ObjConn.execute(sqlstr)
				if cdbl(rs(0))=0 then
					isInsert = true
				end if
				rs.close

				if isInsert then


					sqlstr = "select max(unique_id) from sms_result"
					set rs = objconn.execute(sqlstr)
					if not rs.eof then
						unique_id = cdbl("0"&rs(0)) + 1
					else
						unique_id =1
					end if
					rs.close

					sqlstr = "insert into SMS_Result (unique_id " & vbcrlf 
					sqlstr = sqlstr & ",group_id " & vbcrlf 
					sqlstr = sqlstr & ",cell_no " & vbcrlf 
					sqlstr = sqlstr & ",[description] " & vbcrlf 
					sqlstr = sqlstr & ",send_time " & vbcrlf 
					sqlstr = sqlstr & ",sent_result " & vbcrlf 
					sqlstr = sqlstr & ",people_name " & vbcrlf 
					sqlstr = sqlstr & ",people " & vbcrlf 
					sqlstr = sqlstr & ",postdate " & vbcrlf 
					sqlstr = sqlstr & ",account_name " & vbcrlf 
					sqlstr = sqlstr & ",account_id " & vbcrlf 
					sqlstr = sqlstr & ",system_id " & vbcrlf 
					sqlstr = sqlstr & ",delta_day " & vbcrlf 
					sqlstr = sqlstr & ",delta_hour " & vbcrlf 
					sqlstr = sqlstr & ",delta_minute " & vbcrlf 
					sqlstr = sqlstr & ",visible_flag " & vbcrlf 
					sqlstr = sqlstr & ") values ("&det_unique_id&" " & vbcrlf 
					sqlstr = sqlstr & ","&FnSQL(group_id,0)&" " & vbcrlf 
					sqlstr = sqlstr & ","&FnSQL(cell_no,0)&" " & vbcrlf 
					sqlstr = sqlstr & ","&FnSQL(sms_body,0)&" " & vbcrlf 
					sqlstr = sqlstr & ","&FnSQL(send_time,2)&" " & vbcrlf 
					sqlstr = sqlstr & ",0 " & vbcrlf 
					sqlstr = sqlstr & ","&FnSQL(account_name,0)&" " & vbcrlf 
					sqlstr = sqlstr & ","&FnSQL(account_id,0)&" " & vbcrlf 
					sqlstr = sqlstr & ",getdate() " & vbcrlf 
					sqlstr = sqlstr & ","&FnSQL(session("username"),0)&" " & vbcrlf 
					sqlstr = sqlstr & ","&FnSQL(session("userid"),0)&" " & vbcrlf 
					sqlstr = sqlstr & ",1 " & vbcrlf 						
					sqlstr = sqlstr & ",0 " & vbcrlf 
					sqlstr = sqlstr & ",0 " & vbcrlf 
					sqlstr = sqlstr & ",0 " & vbcrlf 
					sqlstr = sqlstr & ",1 " & vbcrlf 
					sqlstr = sqlstr & ") " & vbcrlf 
					objconn.execute(sqlstr)


				end if

			next
		end if

	else
	
		'選擇媒體
		if isArray(media_ary) then
			for r=0 to ubound(media_ary)
				if trim(media_ary(r)&"")>"" then

					item_str = trim(media_ary(r))
					item_ary = split(item_str,"@")
					isdept = false
					if trim(item_ary(1)&"")="" then
						isdept = true
					end if

					if isdept then
						sqlstr = "select unique_id,[Name],[cell] from members1 where cell>'' and len(cell)=10	and dept_id='"&trim(item_ary(0))&"'"
					else 
						sqlstr = "select unique_id,[Name],[cell] from members1 where cell>'' and len(cell)=10	and unique_id='"&trim(item_ary(0))&"'"
					end if
					set rs = objconn.execute(sqlstr)
					if not rs.eof then
						memb_ary = rs.getrows()
					else 
						if isArray(memb_ary) then
							set memb_ary = nothing
						end if
					end if
					rs.close
					if IsArray(memb_ary) then
						for rows=0 to ubound(memb_ary,2)
							account_id = trim(memb_ary(0,rows)&"")
							account_name = trim(memb_ary(1,rows)&"")
							cell_no = trim(memb_ary(2,rows)&"")

							sqlstr = "select max(unique_id) from sms_result "
							set rs = objconn.execute(sqlstr)
							if not rs.eof then
								det_unique_id = cdbl("0"&rs(0)) + 1
							else
								det_unique_id =1
							end if
							rs.close

							isInsert = false

							sqlstr = "select count(*) from sms_result where group_id='"&group_id&"' and people='"&account_id&"'"
							set rs = ObjConn.execute(sqlstr)
							if cdbl(rs(0))=0 then
								isInsert = true
							end if
							rs.close

							if isInsert then


								sqlstr = "select max(unique_id) from sms_result"
								set rs = objconn.execute(sqlstr)
								if not rs.eof then
									unique_id = cdbl("0"&rs(0)) + 1
								else
									unique_id =1
								end if
								rs.close

								sqlstr = "insert into SMS_Result (unique_id " & vbcrlf 
								sqlstr = sqlstr & ",group_id " & vbcrlf 
								sqlstr = sqlstr & ",cell_no " & vbcrlf 
								sqlstr = sqlstr & ",[description] " & vbcrlf 
								sqlstr = sqlstr & ",send_time " & vbcrlf 
								sqlstr = sqlstr & ",sent_result " & vbcrlf 
								sqlstr = sqlstr & ",people_name " & vbcrlf 
								sqlstr = sqlstr & ",people " & vbcrlf 
								sqlstr = sqlstr & ",postdate " & vbcrlf 
								sqlstr = sqlstr & ",account_name " & vbcrlf 
								sqlstr = sqlstr & ",account_id " & vbcrlf 
								sqlstr = sqlstr & ",system_id " & vbcrlf 
								sqlstr = sqlstr & ",delta_day " & vbcrlf 
								sqlstr = sqlstr & ",delta_hour " & vbcrlf 
								sqlstr = sqlstr & ",delta_minute " & vbcrlf 
								sqlstr = sqlstr & ",visible_flag " & vbcrlf 
								sqlstr = sqlstr & ") values ("&det_unique_id&" " & vbcrlf 
								sqlstr = sqlstr & ","&FnSQL(group_id,0)&" " & vbcrlf 
								sqlstr = sqlstr & ","&FnSQL(cell_no,0)&" " & vbcrlf 
								sqlstr = sqlstr & ","&FnSQL(sms_body,0)&" " & vbcrlf 
								sqlstr = sqlstr & ","&FnSQL(send_time,2)&" " & vbcrlf 
								sqlstr = sqlstr & ",0 " & vbcrlf 
								sqlstr = sqlstr & ","&FnSQL(account_name,0)&" " & vbcrlf 
								sqlstr = sqlstr & ","&FnSQL(account_id,0)&" " & vbcrlf 
								sqlstr = sqlstr & ",getdate() " & vbcrlf 
								sqlstr = sqlstr & ","&FnSQL(session("username"),0)&" " & vbcrlf 
								sqlstr = sqlstr & ","&FnSQL(session("userid"),0)&" " & vbcrlf 
								sqlstr = sqlstr & ",1 " & vbcrlf 						
								sqlstr = sqlstr & ",0 " & vbcrlf 
								sqlstr = sqlstr & ",0 " & vbcrlf 
								sqlstr = sqlstr & ",0 " & vbcrlf 
								sqlstr = sqlstr & ",1 " & vbcrlf 
								sqlstr = sqlstr & ") " & vbcrlf 
								objconn.execute(sqlstr)


							end if

						next
					end if


				end if
			next
		end if



	end if




	sqlstr = "select max(unique_id) from sms"
	set rs = objconn.execute(sqlstr)
	if not rs.eof then
		unique_id = cdbl("0"&rs(0)) + 1
	else
		unique_id =1
	end if
	rs.close

	sqlstr = "insert into SMS (unique_id " & vbcrlf 
	sqlstr = sqlstr & ",group_id " & vbcrlf 
	sqlstr = sqlstr & ",account_name " & vbcrlf 
	sqlstr = sqlstr & ",account_id " & vbcrlf 
	sqlstr = sqlstr & ",[description] " & vbcrlf 
	sqlstr = sqlstr & ",postdate " & vbcrlf 
	sqlstr = sqlstr & ",people_value " & vbcrlf 
	sqlstr = sqlstr & ",people_prompt " & vbcrlf 
	sqlstr = sqlstr & ",people_value1 " & vbcrlf 
	sqlstr = sqlstr & ",people_prompt1 " & vbcrlf 
	sqlstr = sqlstr & ",send_year " & vbcrlf 
	sqlstr = sqlstr & ",send_month " & vbcrlf 
	sqlstr = sqlstr & ",send_day " & vbcrlf 
	sqlstr = sqlstr & ",send_hour " & vbcrlf 
	sqlstr = sqlstr & ",send_minute " & vbcrlf 
	sqlstr = sqlstr & ",send_second " & vbcrlf 
	sqlstr = sqlstr & ",send_date " & vbcrlf 
	sqlstr = sqlstr & ",delta_day " & vbcrlf 
	sqlstr = sqlstr & ",delta_hour " & vbcrlf 
	sqlstr = sqlstr & ",delta_minute " & vbcrlf 
	sqlstr = sqlstr & ",visible_flag " & vbcrlf 
	sqlstr = sqlstr & ") values ("&unique_id&" " & vbcrlf 
	sqlstr = sqlstr & ","&FnSQL(group_id,0)&" " & vbcrlf 
	sqlstr = sqlstr & ","&FnSQL(session("username"),0)&" " & vbcrlf 
	sqlstr = sqlstr & ","&FnSQL(session("userid"),0)&" " & vbcrlf 
	sqlstr = sqlstr & ","&FnSQL(sms_body,0)&" " & vbcrlf 
	sqlstr = sqlstr & ",getdate() " & vbcrlf 
	sqlstr = sqlstr & ","&FnSQL(topeople,0)&" " & vbcrlf 
	sqlstr = sqlstr & ","&FnSQL(topeople_txt,0)&" " & vbcrlf 
	sqlstr = sqlstr & ","&FnSQL(tomedia,0)&" " & vbcrlf 
	sqlstr = sqlstr & ","&FnSQL(tomedia_txt,0)&" " & vbcrlf 
	sqlstr = sqlstr & ","&FnSQL(send_y,1)&" " & vbcrlf 
	sqlstr = sqlstr & ","&FnSQL(send_m,1)&" " & vbcrlf 
	sqlstr = sqlstr & ","&FnSQL(send_d,1)&" " & vbcrlf 
	sqlstr = sqlstr & ","&FnSQL(send_h,1)&" " & vbcrlf 
	sqlstr = sqlstr & ","&FnSQL(send_n,1)&" " & vbcrlf 
	sqlstr = sqlstr & ","&FnSQL(send_s,1)&" " & vbcrlf 
	sqlstr = sqlstr & ","&FnSQL(send_time,2)&" " & vbcrlf 
	sqlstr = sqlstr & ",0 " & vbcrlf 
	sqlstr = sqlstr & ",0 " & vbcrlf 
	sqlstr = sqlstr & ",0 " & vbcrlf 
	sqlstr = sqlstr & ",1 " & vbcrlf 
	sqlstr = sqlstr & ") " & vbcrlf 

	objconn.execute(sqlstr)

	'程式無誤才寫入資料庫
	IF err.Number<>0 Then
		status = "9999"
		status_desc "資料寫入異常。"
		Objconn.RollbackTrans
		'response.end
	ELSE
		ObjConn.CommitTrans
	ENd IF

end if

Dim jsa
Set jsa = jsObject()

jsa("status") = status
jsa("status_desc") = status_desc
jsa("group_id") = group_id

jsa.Flush
response.end%>