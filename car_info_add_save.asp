<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	u_type = request.form("u_type")
	car_no = request.form("car_no")
	car_name = request.form("car_name")
	oil_kind = request.form("oil_kind")
	car_owner = request.form("car_owner")
	buy_gubun = request.form("buy_gubun")
	car_reg_date = request.form("car_reg_date")
	owner_emp_no = request.form("owner_emp_no")	
	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

	if	u_type = "U" then
		sql = "Update overtime set work_date='"&work_date&"',company ='"&company&"',dept='"&dept&"',work_item='"&work_item& _ 
		"',from_time='"&from_time&"',to_time='"&to_time&"',work_gubun='"&work_gubun&"',overtime_amt='"&overtime_amt& _
		"',work_memo='"&work_memo&"',cancel_sw='"&cancel_sw&"',mod_id='"&user_id&"',mod_date=now() where work_date" & _
		" = '"&old_date&"' and mg_ce_id = '"&mg_ce_id&"'"
		dbconn.execute(sql)
	  else
		sql="insert into car_info (car_no,car_name,oil_kind,car_owner,buy_gubun,car_reg_date,owner_emp_no,start_date,last_km,reg_emp_no"& _
		",reg_emp_name,reg_date) values ('"&car_no&"','"&car_name&"','"&oil_kind&"','"&car_owner&"','"&buy_gubun&"','"&car_reg_date& _
		"','"&owner_emp_no&"','"&car_reg_date&"',0,'"&emp_no&"','"&user_name&"',now())"
		dbconn.execute(sql)
	end if
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "변경중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + "변경되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"parent.opener.location.reload();"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
	

%>
