<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

emp_user = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

	u_type = request.form("u_type")
	
	car_no = request.form("car_no")
	car_name = request.form("car_name")
	car_year = request.form("car_year")
	car_reg_date = request.form("car_reg_date")
	owner_emp_name = request.form("owner_emp_name")
	owner_emp_no = request.form("owner_emp_no")
	car_use_dept = request.form("car_use_dept")
	oil_kind = request.form("oil_kind")
	car_owner = request.form("car_owner")
	
	as_car_no = request.form("car_no")
	as_date = request.form("as_date")

	as_cause = request.form("as_cause")
    as_solution = request.form("as_solution")
	as_amount = int(request.form("as_amount"))
	as_amount_sign = request.form("as_amount_sign")
	as_repair_pre_yn = request.form("as_repair_pre_yn")
    as_car_name = request.form("car_name")
    as_owner_emp_no = request.form("owner_emp_no")
    as_owner_emp_name = request.form("owner_emp_name")
    as_use_org_name = request.form("car_use_dept")
	
	set dbconn = server.CreateObject("adodb.connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
    Set Rs_as = Server.CreateObject("ADODB.Recordset")
	Set Rs_emp = Server.CreateObject("ADODB.Recordset")
	Set Rs_trans = Server.CreateObject("ADODB.Recordset")
	dbconn.open dbconnect

        sql="select * from emp_master where emp_no='" + owner_emp_no + "'"
		Set Rs_emp = DbConn.Execute(SQL)
		if not Rs_emp.eof then
			emp_company = Rs_emp("emp_company")
			emp_bonbu = Rs_emp("emp_bonbu")
			emp_saupbu = Rs_emp("emp_saupbu")
			emp_team = Rs_emp("emp_team")
			emp_org_name = Rs_emp("emp_org_name")
			emp_reside_place = Rs_emp("emp_reside_place")
		  else
			emp_company = ""
			emp_bonbu = ""
			emp_saupbu = ""
			emp_team = ""
			emp_org_name = ""
			emp_reside_place = ""
		end if	 
		Rs_emp.close()
		
		sql="select max(as_seq) as max_seq from car_as where as_car_no='" + as_car_no + "' and as_date='" + as_date + "'"
		set Rs_as=dbconn.execute(sql)
		if	isnull(Rs_as("max_seq"))  then
			as_seq = "001"
		  else
			max_seq = "00" + cstr((int(Rs_as("max_seq")) + 1))
			as_seq = right(max_seq,3)
		end if	 
		Rs_as.close()
		
		sql="select max(run_seq) as max_seq from transit_cost where mg_ce_id='" + owner_emp_no + "' and run_date='" + as_date + "'"
		set Rs_trans=dbconn.execute(sql)
		if	isnull(Rs_trans("max_seq"))  then
			run_seq = "01"
		  else
			max_seq = "00" + cstr((int(Rs_trans("max_seq")) + 1))
			run_seq = right(max_seq,2)
		end if	 
		Rs_trans.close() 

	dbconn.BeginTrans

'	if	u_type = "U" then
'		sql = "Update car_as set as_cause='"&as_cause&"',as_solution='"&as_solution&"',as_amount='"&as_amount&"',as_amount_sign='"&as_amount_sign&"',as_repair_pre_yn='"&as_repair_pre_yn&"' where as_car_no = '"&as_car_no&"' and as_date = '"&as_date&"' and as_seq = '"&as_seq&"'"
'		dbconn.execute(sql)
				
'	  else
  
		sql="insert into car_as (as_car_no,as_date,as_seq,as_cause,as_solution,as_amount,as_amount_sign,as_repair_pre_yn,as_car_name,as_owner_emp_no,as_owner_emp_name,as_use_org_name,as_reg_date,as_reg_user) values ('"&as_car_no&"','"&as_date&"','"&as_seq&"','"&as_cause&"','"&as_solution&"','"&as_amount&"','"&as_amount_sign&"','"&as_repair_pre_yn&"','"&as_car_name&"','"&as_owner_emp_no&"','"&as_owner_emp_name&"','"&as_use_org_name&"',now(),'"&emp_user&"')"
		dbconn.execute(sql)
		
		if as_amount_sign = "현금" then
		
		      sql="insert into transit_cost (mg_ce_id,run_date,run_seq,emp_company,bonbu,saupbu,team,org_name,reside_place,car_no,car_name,car_owner,oil_kind,repair_pay,repair_cost,cancel_yn,end_yn,reg_id,reg_user,reg_date,repair_pre_yn,run_memo,company) values ('"&owner_emp_no&"','"&as_date&"','"&run_seq&"','"&emp_company&"','"&emp_bonbu&"','"&emp_saupbu&"','"&emp_team&"','"&emp_org_name&"','"&emp_reside_place&"','"&as_car_no&"','"&as_car_name&"','"&car_owner&"','"&oil_kind&"','"&as_amount_sign&"','"&as_amount&"','N','N','"&emp_no&"','"&emp_user&"',now(),'"&as_repair_pre_yn&"','"&as_solution&"','공통')"
			  
		      dbconn.execute(sql)
		end if
		
'	end if
	
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
