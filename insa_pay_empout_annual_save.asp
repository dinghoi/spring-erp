<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	u_type = request.form("u_type")

'지급항목	
	emp_no = request.form("emp_no")
	pmg_yymm = request.form("pmg_yymm")
	pmg_date = request.form("pmg_date")

	pmg_base_pay =int(request.form("yun_pay"))
	pmg_give_total = int(request.form("yun_pay"))
	
'공제항목
    de_epi_amt = int(request.form("epi_amt"))
	de_deduct_total = int(request.form("epi_amt"))		

	set dbconn = server.CreateObject("adodb.connection")
	Set Rs_emp = Server.CreateObject("ADODB.Recordset")
    Set Rs_year = Server.CreateObject("ADODB.Recordset")
    Set Rs_give = Server.CreateObject("ADODB.Recordset")
    Set Rs_dct = Server.CreateObject("ADODB.Recordset")
    Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
	dbconn.open dbconnect

Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
Set rs_emp = DbConn.Execute(SQL)
if not rs_emp.eof then
    	emp_first_date = rs_emp("emp_first_date")
		emp_in_date = rs_emp("emp_in_date")
		emp_end_date = rs_emp("emp_end_date")
		pmg_emp_type = rs_emp("emp_type")
		pmg_grade = rs_emp("emp_grade")
		pmg_position = rs_emp("emp_position")
		pmg_company = rs_emp("emp_company")
		pmg_bonbu = rs_emp("emp_bonbu")
		pmg_saupbu = rs_emp("emp_saupbu")
		pmg_team = rs_emp("emp_team")
		pmg_org_code = rs_emp("emp_org_code")
		pmg_org_name = rs_emp("emp_org_name")
		pmg_reside_place = rs_emp("emp_reside_place")
		pmg_reside_company = rs_emp("emp_reside_company")
		if rs_emp("emp_yuncha_date") = "1900-01-01" or isNull(rs_emp("emp_yuncha_date")) then
                emp_yuncha_date = rs_emp("emp_in_date")
           else 
                emp_yuncha_date = rs_emp("emp_yuncha_date")
        end if
   else
		emp_first_date = ""
		emp_in_date = ""
		emp_end_date = ""
		emp_yuncha_date = ""
		pmg_emp_type = ""
		pmg_grade = ""
		pmg_position = ""
		pmg_company = ""
		pmg_bonbu = ""
		pmg_saupbu = ""
		pmg_team = ""
		pmg_org_code = ""
		pmg_org_name = ""
		pmg_reside_place = ""
		pmg_reside_company = ""
end if
rs_emp.close()

Sql = "SELECT * FROM pay_bank_account where emp_no = '"&emp_no&"'"
Set rs_bnk = DbConn.Execute(SQL)
if not rs_bnk.eof then
           pmg_bank_name = rs_bnk("bank_name")
           pmg_account_no = rs_bnk("account_no")
		   pmg_account_holder = rs_bnk("account_holder")
	   else
           pmg_bank_name = ""
		   pmg_account_no = ""
		   pmg_account_holder = ""
end if
rs_bnk.close()

dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "Update pay_month_give set pmg_base_pay='"&pmg_base_pay&"',pmg_meals_pay ='"&pmg_meals_pay&"',pmg_postage_pay ='"&pmg_postage_pay&"',pmg_re_pay='"&pmg_re_pay&"',pmg_overtime_pay='"&pmg_overtime_pay&"',pmg_car_pay='"&pmg_car_pay&"',pmg_position_pay='"&pmg_position_pay&"',pmg_custom_pay='"&pmg_custom_pay&"',pmg_job_pay='"&pmg_job_pay&"',pmg_job_support='"&pmg_job_support&"',pmg_jisa_pay='"&pmg_jisa_pay&"',pmg_long_pay='"&pmg_long_pay&"',pmg_disabled_pay='"&pmg_disabled_pay&"',pmg_family_pay='"&pmg_family_pay&"',pmg_school_pay='"&pmg_school_pay&"',pmg_qual_pay='"&pmg_qual_pay&"',pmg_tax_yes='"&pmg_tax_yes&"',pmg_tax_no='"&pmg_tax_no&"',pmg_tax_reduced='"&pmg_tax_reduced&"',pmg_give_total='"&pmg_give_total&"',pmg_bank_name='"&pmg_bank_name&"',pmg_account_no='"&pmg_account_no&"',pmg_account_holder='"&pmg_account_holder&"',pmg_mod_user='"&emp_user&"',pmg_mod_date=now() where pmg_yymm = '"&pmg_yymm&"' and pmg_id = '4' and pmg_emp_no = '"&emp_no&"' and pmg_company = '"&pmg_company&"'"
		dbconn.execute(sql)
		
		sql = "Update pay_month_deduct set de_nps_amt='"&de_nps_amt&"',de_nhis_amt ='"&de_nhis_amt&"',de_epi_amt ='"&de_epi_amt&"',de_longcare_amt ='"&de_longcare_amt&"',de_income_tax='"&de_income_tax&"',de_wetax='"&de_wetax&"',de_other_amt1='"&de_other_amt1&"',de_saving_amt='"&de_saving_amt&"',de_sawo_amt='"&de_sawo_amt&"',de_johab_amt='"&de_johab_amt&"',de_hyubjo_amt='"&de_hyubjo_amt&"',de_school_amt='"&de_school_amt&"',de_nhis_bla_amt='"&de_nhis_bla_amt&"',de_long_bla_amt='"&de_long_bla_amt&"',de_deduct_total='"&de_deduct_total&"',de_mod_user='"&emp_user&"',de_mod_date=now() where de_yymm = '"&pmg_yymm&"' and de_id = '4' and de_emp_no = '"&emp_no&"' and de_company = '"&pmg_company&"'"
		dbconn.execute(sql)
		
	  else
		sql="insert into pay_month_give (pmg_yymm,pmg_id,pmg_emp_no,pmg_company,pmg_date,pmg_emp_name,pmg_emp_type,pmg_org_code,pmg_org_name,pmg_bonbu,pmg_saupbu,pmg_team,pmg_reside_place,pmg_reside_company,pmg_grade,pmg_position,pmg_base_pay,pmg_give_total,pmg_bank_name,pmg_account_no,pmg_account_holder,pmg_reg_date,pmg_reg_user) values ('"&pmg_yymm&"','4','"&emp_no&"','"&pmg_company&"','"&pmg_date&"','"&pmg_emp_name&"','"&pmg_emp_type&"','"&pmg_org_code&"','"&pmg_org_name&"','"&pmg_bonbu&"','"&pmg_saupbu&"','"&pmg_team&"','"&pmg_reside_place&"','"&pmg_reside_company&"','"&pmg_grade&"','"&pmg_position&"','"&pmg_base_pay&"','"&pmg_give_total&"','"&pmg_bank_name&"','"&pmg_account_no&"','"&pmg_account_holder&"',now(),'"&emp_user&"')"
		dbconn.execute(sql)
		
		sql="insert into pay_month_deduct (de_yymm,de_id,de_emp_no,de_company,de_date,de_emp_name,de_emp_type,de_org_code,de_org_name,de_bonbu,de_saupbu,de_team,de_reside_place,de_reside_company,de_grade,de_position,de_epi_amt,de_deduct_total,de_reg_date,de_reg_user) values ('"&pmg_yymm&"','4','"&emp_no&"','"&pmg_company&"','"&pmg_date&"','"&pmg_emp_name&"','"&pmg_emp_type&"','"&pmg_org_code&"','"&pmg_org_name&"','"&pmg_bonbu&"','"&pmg_saupbu&"','"&pmg_team&"','"&pmg_reside_place&"','"&pmg_reside_company&"','"&pmg_grade&"','"&pmg_position&"','"&de_epi_amt&"','"&de_deduct_total&"',now(),'"&emp_user&"')"
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
