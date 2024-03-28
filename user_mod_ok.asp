<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	pass = request.form("pass")
	mod_pass = request.form("mod_re_pass")
	hp = request.form("hp")
	car_yn = request.form("car_yn")
	old_car_yn = request.form("old_car_yn")
	car_no = request.form("car_no")
	old_car_no = request.form("old_car_no")
	car_name = request.form("car_name")
	car_owner = request.form("car_owner")
	oil_kind = request.form("oil_kind")
	curr_date = mid(now(),1,10)

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open DbConnect
	if mod_pass = "" then
		sql = "Update memb set hp ='"&hp&"',car_yn ='"&car_yn&"',mod_id ='"&user_id&"',mod_date=now() where user_id='"&user_id&"'"
	  else
		sql = "Update memb set pass='"&mod_pass&"',hp ='"&hp&"',car_yn ='"&car_yn&"',mod_id ='"&user_id&"',mod_date=now() where user_id='"&user_id&"'"
	end if
	dbconn.execute(sql)

	if car_yn = "Y" then
		sql = "select * from car_info where owner_emp_no ='"&user_id&"'"
		Set rs_car=dbconn.execute(Sql)
		if rs_car.eof or rs_car.bof then
			sql="insert into car_info (car_no,car_name,oil_kind,car_owner,buy_gubun,car_reg_date,owner_emp_no,owner_emp_name,start_date,last_km,reg_emp_no,reg_emp_name,reg_date,insurance_amt) values ('"&car_no&"','"&car_name&"','"&oil_kind&"','개인','구매','"&curr_date&"','"&user_id&"','"&user_name&"','"&curr_date&"',0,'"&user_id&"','"&user_name&"',now(),0)"
			dbconn.execute(sql)
		  else
			if old_car_no = car_no then
				sql = "Update car_info set car_name ='"&car_name&"',oil_kind='"&oil_kind& _ 
				"',mod_emp_no='"&user_id&"',mod_emp_name='"&user_name&"',mod_date=now() where owner_emp_no = '"&user_id&"'"
				dbconn.execute(sql)
			  else
				sql="delete from car_info where owner_emp_no = '"&user_id&"'"
				dbconn.execute(sql)
				sql="insert into car_info (car_no,car_name,oil_kind,car_owner,buy_gubun,car_reg_date,owner_emp_no,start_date,last_km"& _
				",reg_emp_no,reg_emp_name,reg_date,insurance_amt) values ('"&car_no&"','"&car_name&"','"&oil_kind&"','개인','구매','"&curr_date& _
				"','"&user_id&"','"&curr_date&"',0,'"&user_id&"','"&user_name&"',now(),0)"
				dbconn.execute(sql)			  		
			end if
		end if			
	end if

	if car_yn = "N" and old_car_yn = "Y" and car_owner = "개인" then 
		sql="delete from car_info where owner_emp_no = '"&user_id&"'"
		dbconn.execute(sql)
	end if
	
	response.write"<script language=javascript>"
	response.write"alert('변경되었습니다....');"		
	response.write"parent.opener.location.reload();"
	response.write"self.close() ;"
	response.write"</script>"
	
	Response.End
	dbconn.Close()
	Set dbconn = Nothing

%>
	