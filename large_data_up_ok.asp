<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<% 
'	on Error resume next

	company = request.form("company")
	as_type = request.form("as_type")
	request_date = request.form("request_date")
	end_date = request.form("end_date")
	paper_no = request.form("paper_no")
	objFile = request.form("objFile")

'	objFile = SERVER.MapPath(".") & "\srv_upload\주소록.xls"
	set cn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")

	Set DbConn = Server.CreateObject("ADODB.Connection")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	DbConn.Open dbconnect
	
	dbconn.BeginTrans

	cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
	rs.Open "select * from [1:10000]",cn,"0"
	
	rowcount=-1
	xgr = rs.getrows
	rowcount = ubound(xgr,2)
	fldcount = rs.fields.count

	tot_cnt = rowcount + 1
    if rowcount > -1 then
		for i=0 to rowcount
' 구군
		sql_etc = "select * from ce_area where sido = '" + xgr(8,i) +"' and gugun = '" + xgr(9,i) + "'"
		set rs_etc=dbconn.execute(sql_etc)				
		if rs_etc.eof then
			tot_gugun = tot_gugun + 1
			tot_err = tot_err + 1
			mg_ce_id = ""
		  else
			mg_ce_id = rs_etc("mg_ce_id")	  
		end if

		if as_type = "랜공사" or as_type = "이전랜공사" then
			dev_inst_cnt = 0
			ran_cnt = xgr(14,i)
			work_man_cnt = 1
			alba_cnt = 0
		  else
			dev_inst_cnt = xgr(14,i)
			ran_cnt = 0
			work_man_cnt = 1
			alba_cnt = 0
		end if
' CE
		sql_etc = "select * from memb where user_id = '" + mg_ce_id + "'"
		set rs_etc=dbconn.execute(sql_etc)				
		if rs_etc.eof then
			mg_ce = "미등록"
			reside = rs_etc("미등록")
			reside_place = rs_etc("미등록")
			team = rs_etc("미등록")
		  else
			mg_ce = rs_etc("user_name")
			reside = rs_etc("reside")
			reside_place = rs_etc("reside_place")
			team = rs_etc("team")
		end if

' CE
		if (xgr(12,i) = "" or isnull(xgr(12,i))) and (xgr(13,i) = "" or isnull(xgr(12,i))) then
			sql_etc = "select * from memb where user_id = '" + mg_ce_id + "'"
			set rs_etc=dbconn.execute(sql_etc)				
			if rs_etc.eof then
				mg_ce = "미등록"
				reside = "미등록"
				reside_place = "미등록"
				reside_company = "미등록"
				team = "미등록"
			  else
				mg_ce = rs_etc("user_name")
				reside = rs_etc("reside")
				reside_place = rs_etc("reside_place")
				reside_company = rs_etc("reside_company")
				team = rs_etc("team")
			end if
		end if

		if xgr(12,i) <> "" then
			sql_etc = "select * from memb where user_id = '" + cstr(xgr(12,i)) + "'"
			set rs_etc=dbconn.execute(sql_etc)				
			if rs_etc.eof then
				mg_ce = "미등록"
				reside = "미등록"
				reside_place = "미등록"
				reside_company = "미등록"
				team = "미등록"
				mg_ce_id = "미등록"
			  else
				mg_ce = rs_etc("user_name")
				reside = rs_etc("reside")
				reside_place = rs_etc("reside_place")
				reside_company = rs_etc("reside_company")
				team = rs_etc("team")
				mg_ce_id = rs_etc("user_id")
			end if
		end if
							
		if xgr(13,i) <> "" then
			sql_etc = "select * from memb where user_name = '" + xgr(13,i) + "'"
			set rs_etc=dbconn.execute(sql_etc)				
			if rs_etc.eof then
				tot_ce = tot_ce + 1
				tot_err = tot_err + 1
				mg_ce = "미등록"
				reside = "미등록"
				reside_place = "미등록"
				reside_company = "미등록"
				team = "미등록"
				mg_ce_id = "미등록"
			  else
				mg_ce = rs_etc("user_name")
				reside = rs_etc("reside")
				reside_place = rs_etc("reside_place")
				reside_company = rs_etc("reside_company")
				team = rs_etc("team")
				mg_ce_id = rs_etc("user_id")
			end if
		end if


		sql="insert into large_acpt (paper_no,acpt_man,acpt_grade,acpt_user,tel_ddd,tel_no1,tel_no2,hp_ddd,hp_no1,hp_no2,company,dept,sido"& _
		",gugun,dong,addr,mg_ce_id,mg_ce,mg_group,as_memo,request_date,request_time,end_date,as_process,as_type,maker,as_device"& _
		",dev_inst_cnt,ran_cnt,work_man_cnt,alba_cnt,team,reside,reside_place,reside_company,sms,upload_ok,reg_id) values "& _
		"('"&paper_no&"','"&user_name&"','"&user_grade&"','"&xgr(1,i)&"','"&xgr(2,i)&"','"&xgr(3,i)&"','"&xgr(4,i)&"','"&xgr(5,i)& _
		"','"&xgr(6,i)&"','"&xgr(7,i)&"','"&company&"','"&xgr(0,i)&"','"&xgr(8,i)&"','"&xgr(9,i)&"','"&xgr(10,i)&"','"&xgr(11,i)& _
		"','"&mg_ce_id&"','"&mg_ce&"','"&mg_group&"','"&as_type&"','"&request_date&"','1000','"&end_date&"','접수','"&as_type&"','.','.'"& _
		","&dev_inst_cnt&","&ran_cnt&","&work_man_cnt&","&alba_cnt&",'"&team&"','"&reside&"','"&reside_place&"','"&reside_company& _
		"','N','N','"&user_id&"')"
		dbconn.execute(sql)
		next
	end if

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "변경중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = cstr(w_cnt) +" 건 등록 완료되었습니다...."
	end if

	err_msg = cstr(rowcount+1) + " 건 처리되었습니다..."
	response.write"<script language=javascript>"
	response.write"alert('"&err_msg&"');"
	response.write"location.replace('large_data_up.asp');"
	response.write"</script>"
	Response.End

	rs.close
	cn.close
	rs_etc.close
	set rs = nothing
	set cn = nothing
	set rs_etc = nothing
%>