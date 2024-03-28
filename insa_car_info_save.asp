<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	u_type = request.form("u_type")

	car_no = request.form("car_no")
	car_old_no = request.form("car_old_no")

	car_name = request.form("car_name")
	car_year = request.form("car_year")
	oil_kind = request.form("oil_kind")
	car_owner = request.form("car_owner")
	buy_gubun = request.form("buy_gubun")
	rental_company = request.form("rental_company")
	car_company = request.form("car_company")
	car_reg_date = request.form("car_reg_date")
	car_use = request.form("car_use")
	car_use_dept = request.form("car_use_dept")
	owner_emp_no = request.form("owner_emp_no")
	owner_emp_name = request.form("emp_name")
	emp_grade = request.form("emp_grade")
	start_date = request.form("car_reg_date")
	car_status = request.form("car_status")
    car_comment = request.form("car_comment")
	last_km = int(request.form("last_km"))
    last_check_date = request.form("last_check_date")
	end_date = request.form("end_date")
	if car_reg_date = "" or isnull(car_reg_date) then
	   car_reg_date = "1900-01-01"
	end if
	if start_date = "" or isnull(start_date) then
	   start_date = "1900-01-01"
	end if
	if last_check_date = "" or isnull(last_check_date) then
	   last_check_date = "1900-01-01"
	end if
	if end_date = "" or isnull(end_date) then
	   end_date = "1900-01-01"
	end if
	if car_year = "" or isnull(car_year) then
	   car_year = "1900-01-01"
	end if

	'start_time = cstr(start_hh) + cstr(start_mm)

	emp_company = request.form("emp_company")
	emp_org_code = request.form("emp_org_code")
	emp_org_name = request.form("emp_org_name")

	insurance_company = ""
    insurance_date = ""
    insurance_amt = 0

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
	    sql = " delete from car_info " & _
	            "  where car_no ='"&car_old_no&"'"

	    dbconn.execute(sql)

'		sql = "Update car_info set car_year='"&car_year&"',oil_kind ='"&oil_kind&"',car_owner ='"&car_owner&"',rental_company='"&rental_company&"',car_company='"&car_company&"',car_reg_date='"&car_reg_date&"',car_use_dept='"&car_use_dept&"',car_use='"&car_use&"',start_date='"&start_date&"',car_status='"&car_status&"',car_comment='"&car_comment&"',last_km='"&last_km&"',last_check_date='"&last_check_date&"',mod_emp_name='"&emp_user&"',mod_date=now() where car_no = '"&car_no&"'"
'		dbconn.execute(sql)

			sql="insert into car_info (car_no,car_name,car_year,oil_kind,insurance_amt,car_owner,buy_gubun,rental_company,car_reg_date,car_use_dept,car_company,car_use,owner_emp_no,owner_emp_name,start_date,end_date,last_km,last_check_date,car_status,car_comment,reg_emp_name,reg_date) values ('"&car_no&"','"&car_name&"','"&car_year&"','"&oil_kind&"',0,'"&car_owner&"','"&buy_gubun&"','"&rental_company&"','"&car_reg_date&"','"&car_use_dept&"','"&car_company&"','"&car_use&"','"&owner_emp_no&"','"&owner_emp_name&"','"&start_date&"','"&end_date&"','"&last_km&"','"&last_check_date&"','"&car_status&"','"&car_comment&"','"&emp_user&"',now())"
			dbconn.execute(sql)
		Else
			'//기등록 여부 체크
			Dim nCarCnt : nCarCnt = 0
			sql = " select count(1) as cnt from car_info " & _
				"  where car_no ='"&car_no&"'"

			Set rs_car = dbconn.execute(sql)
			If Not(rs_car.bof Or rs_car.eof) Then
				nCarCnt = CInt(rs_car("cnt"))
				end_msg = "이미 등록된 차량입니다."
			End IF
			rs_car.close()
			Set rs_car = Nothing

			If nCarCnt>0 Then
				response.write"<script language=javascript>"
				response.write"alert('"&end_msg&"');"
				response.write"history.back();"
				response.write"</script>"
			End IF


			sql="insert into car_info (car_no,car_name,car_year,oil_kind,insurance_amt,car_owner,buy_gubun,rental_company,car_reg_date,car_use_dept,car_company,car_use,owner_emp_no,owner_emp_name,start_date,last_km,last_check_date,car_status,car_comment,reg_emp_name,reg_date) values ('"&car_no&"','"&car_name&"','"&car_year&"','"&oil_kind&"',0,'"&car_owner&"','"&buy_gubun&"','"&rental_company&"','"&car_reg_date&"','"&car_use_dept&"','"&car_company&"','"&car_use&"','"&owner_emp_no&"','"&owner_emp_name&"','"&start_date&"','"&last_km&"','"&last_check_date&"','"&car_status&"','"&car_comment&"','"&emp_user&"',now())"
			dbconn.execute(sql)

			sql="insert into car_drive_user (use_car_no,use_owner_emp_no,use_date,use_compay,use_org_code,use_org_name,use_emp_name,use_emp_grade,use_reg_date,use_reg_user) values ('"&car_no&"','"&owner_emp_no&"','"&start_date&"','"&emp_company&"','"&emp_org_code&"','"&emp_org_name&"','"&owner_emp_name&"','"&emp_grade&"',now(),'"&emp_user&"')"
			dbconn.execute(sql)

		end if

	if Err.number <> 0 then
		dbconn.RollbackTrans
		end_msg = sms_msg + "저장중 Error가 발생하였습니다...."
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
