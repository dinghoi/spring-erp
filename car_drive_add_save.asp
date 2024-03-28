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
	old_run_seq = request.form("run_seq")
	car_no = request.form("car_no")
	car_name = request.form("car_name")
	oil_kind = request.form("oil_kind")
	car_owner = request.form("car_owner")
	last_km = int(request.form("last_km"))
	start_company = request.form("start_company")
	start_point = request.form("start_point")
	start_km = int(request.form("start_km"))
	start_hh = request.form("start_hh")
	start_mm = request.form("start_mm")
	start_time = cstr(start_hh) + cstr(start_mm)
	end_company = request.form("end_company")
	end_point = request.form("end_point")
	end_km = int(request.form("end_km"))
	end_hh = request.form("end_hh")
	end_mm = request.form("end_mm")
	end_time = cstr(end_hh) + cstr(end_mm)
	run_memo = request.form("run_memo")
	far = int(request.form("far"))
	if car_owner = "개인" then
		somopum = far * 25
	  else
	  	somopum = 0
	end if

	oil_amt = request.form("oil_amt")
	if oil_amt = "" or isnull(oil_amt) then
		oil_amt = 0
	end if
	oil_amt = int(oil_amt)

	oil_pay = request.form("oil_pay")	

	oil_price = request.form("oil_price")
	if oil_price = "" or isnull(oil_price) then
		oil_price = 0
	end if
	oil_price = int(oil_price)
	
	parking_pay = request.form("parking_pay")	

	parking = request.form("parking")
	if parking = "" or isnull(parking) then
		parking = 0
	end if
	parking = int(parking)

	toll_pay = request.form("toll_pay")	

	toll = request.form("toll")
	if toll = "" or isnull(toll) then
		toll = 0
	end if
	toll = int(toll)

	cancel_yn = request.form("cancel_yn")
	end_yn = request.form("end_yn")

	mod_id = request.form("mod_id")
	mod_user = request.form("mod_user")
	mod_date = request.form("mod_date")

	if mod_id <> "" then
		mod_yymmdd = datevalue(mod_date)
		mod_hhmm = formatdatetime(mod_date,4)
		mod_date = cstr(mod_yymmdd) + " " + cstr(mod_hhmm)
	end if

	company = end_company
	if company = "집" or company = "본사(회사)" or company = "기타" or company = "케이원정보통신" then
		company = start_company
	end if
	if company = "집" or company = "본사(회사)" or company = "기타" or company = "케이원정보통신" then
		company = "공통"
	end if
	if isnull(reside_company) then
		reside_company = ""
	end if 
	if company = "공통" and reside_company <> "" then
		company = reside_company
	end if

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	if	u_type = "U" then
		sql = "delete from transit_cost where run_date ='"&old_date&"' and mg_ce_id='"&mg_ce_id&"' and run_seq= '"&run_seq&"'"
		dbconn.execute(sql)
	end if

	sql = "select max(run_seq) as max_seq from transit_cost where mg_ce_id = '"&mg_ce_id&"' and run_date = '"&run_date&"'"
	set rs = dbconn.execute(sql)
	if rs.eof or rs.bof then
		max_seq = 0
	  else  	
		max_seq = rs("max_seq")
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

	if run_date = old_date then
		run_seq = old_run_seq
	end if
	
	sql = "select * from memb where user_id = '"&mg_ce_id&"'"
	set rs_memb=dbconn.execute(sql)		

	if isnull(mod_id) or mod_id = "" then
		sql="insert into transit_cost (mg_ce_id,user_name,user_grade,run_date,run_seq,emp_company,bonbu,saupbu,team,org_name,reside_place,car_no,car_name,car_owner,oil_kind,start_company,start_point,start_km,start_time,end_company,end_point,end_km,end_time,far,run_memo,company,somopum"& _
		",oil_amt,oil_pay,oil_price,parking_pay,parking,toll_pay,toll,cancel_yn,end_yn,reg_id,reg_user,reg_date) values ("& _
		"'"&mg_ce_id&"','"&rs_memb("user_name")&"','"&rs_memb("user_grade")&"','"&run_date&"','"&run_seq&"','"&emp_company&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name&"','"&reside_place& _
		"','"&car_no&"','"&car_name&"','"&car_owner&"','"&oil_kind&"','"&start_company&"','"&start_point&"',"&start_km&",'"&start_time& _
		"','"&end_company&"','"&end_point&"',"&end_km&",'"&end_time&"',"&far&",'"&run_memo&"','"&company&"',"&somopum& _
		","&oil_amt&",'"&oil_pay&"',"&oil_price&",'"&parking_pay&"',"&parking&",'"&toll_pay&"',"&toll&",'"&cancel_yn&"','"&end_yn& _
		"','"&user_id&"','"&user_name&"',now())"
		dbconn.execute(sql)
	  else
		sql="insert into transit_cost (mg_ce_id,user_name,user_grade,run_date,run_seq,emp_company,bonbu,saupbu,team,org_name,reside_place,car_no,car_name,car_owner"& _
		",oil_kind,start_company,start_point,start_km,start_time,end_company,end_point,end_km,end_time,far,run_memo,company,somopum"& _
		",oil_amt,oil_pay,oil_price,parking_pay,parking,toll_pay,toll,cancel_yn,end_yn,reg_id,reg_user,reg_date,mod_id,mod_user"& _
		",mod_date) values ('"&mg_ce_id&"','"&rs_memb("user_name")&"','"&rs_memb("user_grade")&"','"&run_date&"','"&run_seq&"','"&emp_company&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name& _
		"','"&reside_place&"','"&car_no&"','"&car_name&"','"&car_owner&"','"&oil_kind&"','"&start_company&"','"&start_point&"',"&start_km& _
		",'"&start_time&"','"&end_company&"','"&end_point&"',"&end_km&",'"&end_time&"',"&far&",'"&run_memo&"','"&company&"',"&somopum& _
		","&oil_amt&",'"&oil_pay&"',"&oil_price&",'"&parking_pay&"',"&parking&",'"&toll_pay&"',"&toll&",'"&cancel_yn& _
		"','"&end_yn&"','"&user_id&"','"&user_name&"',now(),'"&mod_id&"','"&mod_user&"','"&mod_date&"')"
		dbconn.execute(sql)
	end if
	if end_km > last_km then
		sql = "Update car_info set last_km="&end_km&" where car_no = '"&car_no&"'"
		dbconn.execute(sql)
	end if
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "변경중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "변경되었습니다...."
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
