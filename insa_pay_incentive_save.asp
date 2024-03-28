<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	u_type = request.form("u_type")
	in_pmg_id = request.form("in_pmg_id")

'지급항목	
	emp_no = request.form("emp_no")
	pmg_yymm = request.form("pmg_yymm")
	pmg_date = request.form("pmg_date")
	pmg_in_date = request.form("emp_in_date")
	pmg_emp_name = request.form("pmg_emp_name")
	pmg_emp_type = request.form("pmg_emp_type")
	pmg_grade = request.form("pmg_grade")
	pmg_position = request.form("pmg_position")
	pmg_company = request.form("pmg_company")
	pmg_org_code = request.form("pmg_org_code")
	pmg_org_name = request.form("pmg_org_name")
	pmg_bonbu = request.form("pmg_bonbu")
	pmg_saupbu = request.form("pmg_saupbu")
	pmg_team = request.form("pmg_team")
	pmg_reside_place = request.form("pmg_reside_place")
	pmg_reside_company = request.form("pmg_reside_company")
	
	cost_group = request.form("cost_group")
	cost_center = request.form("cost_center")
	
	pmg_bank_name = request.form("pmg_bank_name")
	pmg_account_no = request.form("pmg_account_no")
	pmg_account_holder = request.form("pmg_account_holder")
	
	pmg_base_pay =int(request.form("pmg_base_pay"))
	pmg_meals_pay = 0
	pmg_postage_pay = 0
	pmg_re_pay = 0
	pmg_overtime_pay = 0
	pmg_car_pay = 0
	pmg_position_pay = 0
	pmg_custom_pay = 0
	pmg_job_pay = 0
	pmg_job_support = 0
	pmg_jisa_pay = 0
	pmg_long_pay = 0
	pmg_disabled_pay = 0
	pmg_family_pay = 0
	pmg_school_pay = 0
	pmg_qual_pay = 0
	pmg_other_pay1 = 0
	pmg_other_pay2 = 0
	pmg_other_pay3 = 0
	pmg_tax_yes = 0
	pmg_tax_no = 0
	pmg_tax_reduced = 0
	pmg_give_total = int(request.form("pmg_give_tot"))
	
'공제항목
    de_nps_amt = 0
    de_nhis_amt = 0
    de_epi_amt = int(request.form("de_epi_amt"))
	de_longcare_amt = 0
    de_income_tax = int(request.form("de_income_tax"))
    de_wetax = int(request.form("de_wetax"))
    de_other_amt1 = 0
    'de_saving_amt = int(request.form("de_saving_amt"))
	de_saving_amt = 0
	de_other_amt2 = 0
	de_other_amt3 = 0
    de_sawo_amt = 0
    'de_johab_amt = int(request.form("de_johab_amt"))
	de_johab_amt = 0
	de_special_tax = 0
    de_hyubjo_amt = 0
    de_school_amt = 0
    de_nhis_bla_amt = 0
    de_long_bla_amt = 0
	de_deduct_total = int(request.form("de_deduct_tot"))		

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "Update pay_month_give set pmg_in_date='"&pmg_in_date&"',cost_group='"&cost_group&"',cost_center='"&cost_center&"',pmg_base_pay='"&pmg_base_pay&"',pmg_meals_pay ='"&pmg_meals_pay&"',pmg_postage_pay ='"&pmg_postage_pay&"',pmg_re_pay='"&pmg_re_pay&"',pmg_overtime_pay='"&pmg_overtime_pay&"',pmg_car_pay='"&pmg_car_pay&"',pmg_position_pay='"&pmg_position_pay&"',pmg_custom_pay='"&pmg_custom_pay&"',pmg_job_pay='"&pmg_job_pay&"',pmg_job_support='"&pmg_job_support&"',pmg_jisa_pay='"&pmg_jisa_pay&"',pmg_long_pay='"&pmg_long_pay&"',pmg_disabled_pay='"&pmg_disabled_pay&"',pmg_family_pay='"&pmg_family_pay&"',pmg_school_pay='"&pmg_school_pay&"',pmg_qual_pay='"&pmg_qual_pay&"',pmg_tax_yes='"&pmg_tax_yes&"',pmg_tax_no='"&pmg_tax_no&"',pmg_tax_reduced='"&pmg_tax_reduced&"',pmg_give_total='"&pmg_give_total&"',pmg_bank_name='"&pmg_bank_name&"',pmg_account_no='"&pmg_account_no&"',pmg_account_holder='"&pmg_account_holder&"',pmg_mod_user='"&emp_user&"',pmg_mod_date=now() where pmg_yymm = '"&pmg_yymm&"' and pmg_id = '"&in_pmg_id&"' and pmg_emp_no = '"&emp_no&"'"
		dbconn.execute(sql)
		
		sql = "Update pay_month_deduct set cost_group='"&cost_group&"',cost_center='"&cost_center&"',de_nps_amt='"&de_nps_amt&"',de_nhis_amt ='"&de_nhis_amt&"',de_epi_amt ='"&de_epi_amt&"',de_longcare_amt ='"&de_longcare_amt&"',de_income_tax='"&de_income_tax&"',de_wetax='"&de_wetax&"',de_other_amt1='"&de_other_amt1&"',de_saving_amt='"&de_saving_amt&"',de_sawo_amt='"&de_sawo_amt&"',de_johab_amt='"&de_johab_amt&"',de_hyubjo_amt='"&de_hyubjo_amt&"',de_school_amt='"&de_school_amt&"',de_nhis_bla_amt='"&de_nhis_bla_amt&"',de_long_bla_amt='"&de_long_bla_amt&"',de_deduct_total='"&de_deduct_total&"',de_mod_user='"&emp_user&"',de_mod_date=now() where de_yymm = '"&pmg_yymm&"' and de_id = '"&in_pmg_id&"' and de_emp_no = '"&emp_no&"'"
		dbconn.execute(sql)
		
	  else
		sql="insert into pay_month_give (pmg_yymm,pmg_id,pmg_emp_no,pmg_company,pmg_date,pmg_in_date,pmg_emp_name,pmg_emp_type,pmg_org_code,pmg_org_name,pmg_bonbu,pmg_saupbu,pmg_team,pmg_reside_place,pmg_reside_company,pmg_grade,pmg_position,pmg_base_pay,pmg_meals_pay,pmg_postage_pay,pmg_re_pay,pmg_overtime_pay,pmg_car_pay,pmg_position_pay,pmg_custom_pay,pmg_job_pay,pmg_job_support,pmg_jisa_pay,pmg_long_pay,pmg_disabled_pay,pmg_family_pay,pmg_school_pay,pmg_qual_pay,pmg_other_pay1,pmg_other_pay2,pmg_other_pay3,pmg_tax_yes,pmg_tax_no,pmg_tax_reduced,pmg_give_total,pmg_bank_name,pmg_account_no,pmg_account_holder,cost_group,cost_center,pmg_reg_date,pmg_reg_user) values ('"&pmg_yymm&"','"&in_pmg_id&"','"&emp_no&"','"&pmg_company&"','"&pmg_date&"','"&pmg_in_date&"','"&pmg_emp_name&"','"&pmg_emp_type&"','"&pmg_org_code&"','"&pmg_org_name&"','"&pmg_bonbu&"','"&pmg_saupbu&"','"&pmg_team&"','"&pmg_reside_place&"','"&pmg_reside_company&"','"&pmg_grade&"','"&pmg_position&"','"&pmg_base_pay&"','"&pmg_meals_pay&"','"&pmg_postage_pay&"','"&pmg_re_pay&"','"&pmg_overtime_pay&"','"&pmg_car_pay&"','"&pmg_position_pay&"','"&pmg_custom_pay&"','"&pmg_job_pay&"','"&pmg_job_support&"','"&pmg_jisa_pay&"','"&pmg_long_pay&"','"&pmg_disabled_pay&"','"&pmg_family_pay&"','"&pmg_school_pay&"','"&pmg_qual_pay&"','"&pmg_other_pay1&"','"&pmg_other_pay2&"','"&pmg_other_pay3&"','"&pmg_tax_yes&"','"&pmg_tax_no&"','"&pmg_tax_reduced&"','"&pmg_give_total&"','"&pmg_bank_name&"','"&pmg_account_no&"','"&pmg_account_holder&"','"&cost_group&"','"&cost_center&"',now(),'"&emp_user&"')"
		
		dbconn.execute(sql)
		
		sql="insert into pay_month_deduct (de_yymm,de_id,de_emp_no,de_company,de_date,de_emp_name,de_emp_type,de_org_code,de_org_name,de_bonbu,de_saupbu,de_team,de_reside_place,de_reside_company,de_grade,de_position,de_nps_amt,de_nhis_amt,de_epi_amt,de_longcare_amt,de_income_tax,de_wetax,de_other_amt1,de_saving_amt,de_sawo_amt,de_johab_amt,de_hyubjo_amt,de_school_amt,de_nhis_bla_amt,de_long_bla_amt,de_deduct_total,cost_group,cost_center,de_reg_date,de_reg_user) values ('"&pmg_yymm&"','"&in_pmg_id&"','"&emp_no&"','"&pmg_company&"','"&pmg_date&"','"&pmg_emp_name&"','"&pmg_emp_type&"','"&pmg_org_code&"','"&pmg_org_name&"','"&pmg_bonbu&"','"&pmg_saupbu&"','"&pmg_team&"','"&pmg_reside_place&"','"&pmg_reside_company&"','"&pmg_grade&"','"&pmg_position&"','"&de_nps_amt&"','"&de_nhis_amt&"','"&de_epi_amt&"','"&de_longcare_amt&"','"&de_income_tax&"','"&de_wetax&"','"&de_other_amt1&"','"&de_saving_amt&"','"&de_sawo_amt&"','"&de_johab_amt&"','"&de_hyubjo_amt&"','"&de_school_amt&"','"&de_nhis_bla_amt&"','"&de_long_bla_amt&"','"&de_deduct_total&"','"&cost_group&"','"&cost_center&"',now(),'"&emp_user&"')"
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
