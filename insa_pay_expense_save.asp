<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	u_type = request.form("u_type")
	ex_tax_id = request.form("ex_tax_id")
	ex_code = request.form("ex_code")

	ex_emp_no = request.form("ex_emp_no")
	rever_yymm = request.form("rever_yymm")
	ex_date = request.form("ex_date")
	ex_emp_name = request.form("ex_emp_name")
	ex_pay_date = request.form("ex_pay_date")
	ex_code_name = request.form("ex_code_name")
	ex_company = request.form("ex_company")
	ex_bonbu = request.form("ex_bonbu")
	ex_saupbu = request.form("ex_saupbu")
	ex_team = request.form("ex_team")
	ex_reside_place = request.form("ex_reside_place")
	ex_reside_company = request.form("ex_reside_company")
	ex_org_name = request.form("ex_org_name")
	ex_comment = request.form("ex_comment")
	
	ex_amount =int(request.form("ex_amount"))
	ex_work_cnt = int(request.form("ex_work_cnt"))

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "Update pay_expense set ex_work_cnt='"&ex_work_cnt&"',ex_amount ='"&ex_amount&"',ex_comment ='"&ex_comment&"',ex_mod_user='"&emp_user&"',ex_mod_date=now() where ex_date = '"&ex_date&"' and ex_emp_no = '"&ex_emp_no&"' and ex_deduct_id = '"&ex_deduct_id&"' and ex_code_name = '"&ex_code_name&"'"
		dbconn.execute(sql)
		
	  else
		sql="insert into pay_expense (ex_date,ex_emp_no,ex_deduct_id,ex_code_name,ex_code,rever_yymm,ex_pay_date,ex_tax_id,ex_emp_name,ex_work_cnt,ex_amount,ex_company,ex_bonbu,ex_saupbu,ex_team,ex_reside_place,ex_reside_company,ex_org_name,ex_comment,ex_reg_date,ex_reg_user) values ('"&ex_date&"','"&ex_emp_no&"','"&ex_deduct_id&"','"&ex_code_name&"','"&ex_code&"','"&rever_yymm&"','"&ex_pay_date&"','"&ex_tax_id&"','"&ex_emp_name&"','"&ex_work_cnt&"','"&ex_amount&"','"&ex_company&"','"&ex_bonbu&"','"&ex_saupbu&"','"&ex_team&"','"&ex_reside_place&"','"&ex_reside_company&"','"&ex_org_name&"','"&ex_comment&"',now(),'"&emp_user&"')"
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
