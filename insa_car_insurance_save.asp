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

	ins_car_no = request.form("car_no")
	ins_date = request.form("ins_date")

	ins_old_date = request.form("ins_old_date")

	ins_amount = int(request.form("ins_amount"))
    ins_company = request.form("ins_company")
    ins_last_date = request.form("ins_last_date")
    ins_man1 = request.form("ins_man1")
    ins_man2 = request.form("ins_man2")
    ins_object = request.form("ins_object")
    ins_self = request.form("ins_self")
    ins_injury = request.form("ins_injury")
    ins_self_car = request.form("ins_self_car")
    ins_age = request.form("ins_age")
    ins_comment = request.form("ins_comment")
	ins_contract_yn = request.form("ins_contract_yn")
	if ins_contract_yn = "N" then
	       ins_comment = "필요시 제안사에서 운영"
	   else
	       ins_comment = ""
    end if
    ins_scramble = request.form("ins_scramble")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then

		sql = " delete from car_insurance " & _
	            "  where ins_car_no ='"&ins_car_no&"' and ins_date = '"&ins_old_date&"'"

	    dbconn.execute(sql)

'		sql = "Update car_insurance set ins_amount='"&ins_amount&"',ins_company ='"&ins_company&"',ins_last_date ='"&ins_last_date&"',ins_man1='"&ins_man1&"',ins_man2='"&ins_man2&"',ins_object='"&ins_object&"',ins_self='"&ins_self&"',ins_injury='"&ins_injury&"',ins_self_car='"&ins_self_car&"',ins_age='"&ins_age&"',ins_comment='"&ins_comment&"',ins_scramble='"&ins_scramble&"' where ins_car_no = '"&ins_car_no&"' and ins_date = '"&ins_date&"'"
'		dbconn.execute(sql)

'		sql = "Update car_info set insurance_company='"&ins_company&"',insurance_date ='"&ins_date&"',insurance_amt ='"&ins_amount&"',mod_emp_name='"&emp_user&"',mod_date=now() where car_no = '"&ins_car_no&"'"
'		dbconn.execute(sql)

'	  else
    end if
		sql="insert into car_insurance (ins_car_no,ins_date,ins_amount,ins_company,ins_last_date,ins_man1,ins_man2,ins_object,ins_self,ins_injury,ins_self_car,ins_age,ins_comment,ins_contract_yn,ins_scramble,ins_reg_date,ins_reg_user) values ('"&ins_car_no&"','"&ins_date&"','"&ins_amount&"','"&ins_company&"','"&ins_last_date&"','"&ins_man1&"','"&ins_man2&"','"&ins_object&"','"&ins_self&"','"&ins_injury&"','"&ins_self_car&"','"&ins_age&"','"&ins_comment&"','"&ins_contract_yn&"','"&ins_scramble&"',now(),'"&emp_user&"')"
		dbconn.execute(sql)

		sql = "Update car_info set insurance_company='"&ins_company&"',insurance_date ='"&ins_last_date&"',insurance_amt ='"&ins_amount&"',mod_emp_name='"&emp_user&"',mod_date=now() where car_no = '"&ins_car_no&"'"
		dbconn.execute(sql)


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
