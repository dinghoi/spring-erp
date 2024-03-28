<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	u_type = request.form("u_type")
	mg_ce_id = request.form("mg_ce_id")
	run_date = request.form("run_date")
	old_date = request.form("old_date")
	run_seq = request.form("run_seq")
	oil_kind = request.form("oil_kind")
	car_owner = "대중교통"
	start_point = request.form("start_point")
	start_hh = request.form("start_hh")
	start_mm = request.form("start_mm")
	start_time = cstr(start_hh) + cstr(start_mm)
	company = request.form("company")
	end_point = request.form("end_point")
	end_hh = request.form("end_hh")
	end_mm = request.form("end_mm")
	end_time = cstr(end_hh) + cstr(end_mm)
	transit = request.form("transit")	
	payment = request.form("payment")
	run_memo = request.form("run_memo")
	fare = int(request.form("fare"))
	end_yn = request.form("end_yn")	
	cancel_yn = request.form("cancel_yn")	

	mod_id = request.form("mod_id")
	mod_user = request.form("mod_user")
	mod_date = request.form("mod_date")

	if mod_id <> "" then
		mod_yymmdd = datevalue(mod_date)
		mod_hhmm = formatdatetime(mod_date,4)
		mod_date = cstr(mod_yymmdd) + " " + cstr(mod_hhmm)
	end if
	
	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	if	u_type = "U" then
		sql = "delete from transit_cost where run_date ='"&old_date&"' and mg_ce_id='"&mg_ce_id&"' and run_seq='"&run_seq&"'"
		dbconn.execute(sql)
	end if

	sql = "select max(run_seq) as max_seq from transit_cost where mg_ce_id = '"&mg_ce_id&"' and run_date = '"&run_date&"'"
	set rs = dbconn.execute(sql)
	if rs.eof or rs.bof then
		max_seq = 0
	  else  	
		max_seq = int(rs("max_seq"))
	end if
	if isnull(rs("max_seq")) then
		max_seq = 0
	end if
	max_seq = max_seq + 1
	if max_seq < 10 then
		run_seq = "0" + cstr(max_seq)
	  else
		run_seq = cstr(max_seq)
	end if
	rs.Close()		
	
	sql = "select * from memb where user_id = '"&mg_ce_id&"'"
	set rs_memb=dbconn.execute(sql)		

	if isnull(mod_id) or mod_id = "" then
		sql="insert into transit_cost (mg_ce_id,user_name,user_grade,run_date,run_seq,emp_company,bonbu,saupbu,team,org_name,reside_place,car_owner,start_point"& _
		",start_km,start_time,end_point,end_km,end_time,transit,payment,fare,run_memo,company,cancel_yn,end_yn,reg_id,reg_user,reg_date) "& _
		"values ('"&mg_ce_id&"','"&rs_memb("user_name")&"','"&rs_memb("user_grade")&"','"&run_date&"','"&run_seq&"','"&emp_company&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name& _
		"','"&reside_place&"','"&car_owner&"','"&start_point&"',0,'"&start_time&"','"&end_point&"',0,'"&end_time&"','"&transit&"','"&payment& _
		"',"&fare&",'"&run_memo&"','"&company&"','"&cancel_yn&"','"&end_yn&"','"&user_id&"','"&user_name&"',now())"
		dbconn.execute(sql)
	  else
		sql="insert into transit_cost (mg_ce_id,user_name,user_grade,run_date,run_seq,emp_company,bonbu,saupbu,team,org_name,reside_place,car_owner,start_point"& _
		",start_km,start_time,end_point,end_km,end_time,transit,payment,fare,run_memo,company,cancel_yn,end_yn,reg_id,reg_user,reg_date,mod_id"& _
		",mod_user,mod_date) values ('"&mg_ce_id&"','"&rs_memb("user_name")&"','"&rs_memb("user_grade")&"','"&run_date&"','"&run_seq&"','"&emp_company&"','"&bonbu&"','"&saupbu&"','"&team& _
		"','"&org_name&"','"&reside_place&"','"&car_owner&"','"&start_point&"',0,'"&start_time&"','"&end_point&"',0,'"&end_time&"','"&transit& _
		"','"&payment&"',"&fare&",'"&run_memo&"','"&company&"','"&cancel_yn&"','"&end_yn&"','"&user_id&"','"&user_name&"',now(),'"&mod_id& _
		"','"&mod_name&"','"&mod_date&"')"
		dbconn.execute(sql)
	end if
		
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "저장중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "저장되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
	

%>
