<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	u_type = request.form("u_type")

	car_no = request.form("car_no")
	car_name = request.form("car_name")
	car_year = request.form("car_year")
	car_reg_date = request.form("car_reg_date")
	old_owner_emp_name = request.form("old_owner_emp_name")
	old_owner_emp_no = request.form("old_owner_emp_no")

	use_car_no = request.form("car_no")
	use_date = request.form("use_date")
	use_owner_emp_no = request.form("owner_emp_no")
	use_emp_name = request.form("emp_name")
	use_emp_grade = request.form("emp_grade")
    use_company = request.form("emp_company")
    use_org_code = request.form("emp_org_code")
    use_org_name = request.form("emp_org_name")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "Update car_drive_user set use_end_date='"&use_date&"' where use_car_no = '"&use_car_no&"' and use_owner_emp_no = '"&old_owner_emp_no&"' and use_date = '"&use_date&"'"
		dbconn.execute(sql)

		sql = "Update car_info set owner_emp_no='"&use_owner_emp_no&"',owner_emp_name ='"&use_emp_name&"',car_use_dept ='"&use_org_name&"',mod_emp_name='"&emp_user&"',mod_date=now() where car_no = '"&use_car_no&"'"
		dbconn.execute(sql)

	  else
		sql="insert into car_drive_user (use_car_no,use_owner_emp_no,use_date,use_company,use_org_code,use_org_name,use_emp_name,use_emp_grade,use_reg_date,use_reg_user) values ('"&use_car_no&"','"&use_owner_emp_no&"','"&use_date&"','"&use_company&"','"&use_org_code&"','"&use_org_name&"','"&use_emp_name&"','"&use_emp_grade&"',now(),'"&emp_user&"')"
		dbconn.execute(sql)

		sql = "Update car_info set owner_emp_no='"&use_owner_emp_no&"',owner_emp_name ='"&use_emp_name&"',car_use_dept ='"&use_org_name&"',mod_emp_name='"&emp_user&"',mod_date=now() where car_no = '"&use_car_no&"'"
		dbconn.execute(sql)


	end if

	if Err.number <> 0 then
		dbconn.RollbackTrans
		end_msg = sms_msg + "자장중 Error가 발생하였습니다...."
	else
		dbconn.CommitTrans
		end_msg = sms_msg + "저장되었습니다...."
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
