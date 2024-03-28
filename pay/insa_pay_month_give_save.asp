<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
On Error Resume Next
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim u_type, pmg_yymm, pmg_date, pmg_in_date, pmg_emp_name, pmg_emp_type
Dim pmg_grade, pmg_position, pmg_company, pmg_org_code, pmg_org_name, pmg_bonbu
Dim pmg_saupbu, pmg_team, pmg_reside_place, pmg_reside_company, cost_group, cost_center
Dim pmg_bank_name, pmg_account_no, pmg_account_holder, pmg_base_pay, pmg_meals_pay
Dim pmg_postage_pay, pmg_re_pay
Dim pmg_overtime_pay, pmg_car_pay, pmg_position_pay, pmg_custom_pay, pmg_job_pay
Dim pmg_job_support, pmg_jisa_pay, pmg_long_pay, pmg_disabled_pay, pmg_family_pay
Dim pmg_school_pay, pmg_qual_pay, pmg_other_pay1, pmg_other_pay2, pmg_other_pay3
Dim pmg_tax_yes, pmg_tax_no, pmg_tax_reduced, pmg_give_total
Dim de_nps_amt, de_nhis_amt, de_epi_amt, de_longcare_amt, de_income_tax
Dim de_wetax, de_other_amt1, de_saving_amt, de_other_amt2, de_other_amt3
Dim de_sawo_amt, de_johab_amt, de_special_tax, de_hyubjo_amt, de_school_amt
Dim de_nhis_bla_amt, de_long_bla_amt, de_year_incom_tax, de_year_wetax
Dim de_year_incom_tax2, de_year_wetax2, de_deduct_total, end_msg

u_type = Request.Form("u_type")

'지급항목
emp_no = Request.Form("emp_no")
pmg_yymm = Request.Form("pmg_yymm")
pmg_date = Request.Form("pmg_date")
pmg_in_date = Request.Form("emp_in_date")
pmg_emp_name = Request.Form("pmg_emp_name")
pmg_emp_type = Request.Form("pmg_emp_type")
pmg_grade = Request.Form("pmg_grade")
pmg_position = Request.Form("pmg_position")
pmg_company = Request.Form("pmg_company")
pmg_org_code = Request.Form("pmg_org_code")
pmg_org_name = Request.Form("pmg_org_name")
pmg_bonbu = Request.Form("pmg_bonbu")
pmg_saupbu = Request.Form("pmg_saupbu")
pmg_team = Request.Form("pmg_team")
pmg_reside_place = Request.Form("pmg_reside_place")
pmg_reside_company = Request.Form("pmg_reside_company")

cost_group = Request.Form("cost_group")
cost_center = Request.Form("cost_center")

pmg_bank_name = Request.Form("pmg_bank_name")
pmg_account_no = Request.Form("pmg_account_no")
pmg_account_holder = Request.Form("pmg_account_holder")

pmg_base_pay = CLng(f_toString(Request.Form("pmg_base_pay"), 0))
pmg_meals_pay = CLng(f_toString(Request.Form("pmg_meals_pay"), 0))
pmg_postage_pay = CLng(f_toString(Request.Form("pmg_postage_pay"), 0))
pmg_re_pay = CLng(f_toString(Request.Form("pmg_re_pay"), 0))
pmg_overtime_pay = CLng(f_toString(Request.Form("pmg_overtime_pay"), 0))
pmg_car_pay = CLng(f_toString(Request.Form("pmg_car_pay"), 0))
pmg_position_pay = CLng(f_toString(Request.Form("pmg_position_pay"), 0))
pmg_custom_pay = CLng(f_toString(Request.Form("pmg_custom_pay"), 0))
pmg_job_pay = CLng(f_toString(Request.Form("pmg_job_pay"), 0))
pmg_job_support = CLng(f_toString(Request.Form("pmg_job_support"), 0))
pmg_jisa_pay = CLng(f_toString(Request.Form("pmg_jisa_pay"), 0))
pmg_long_pay = CLng(f_toString(Request.Form("pmg_long_pay"), 0))
pmg_disabled_pay = CLng(f_toString(Request.Form("pmg_disabled_pay"), 0))
pmg_family_pay = CLng(f_toString(Request.Form("pmg_family_pay"), 0))
pmg_school_pay = CLng(f_toString(Request.Form("pmg_school_pay"), 0))
pmg_qual_pay = CLng(f_toString(Request.Form("pmg_qual_pay"), 0))
pmg_other_pay1 = CLng(f_toString(Request.Form("pmg_other_pay1"), 0))
pmg_other_pay2 = CLng(f_toString(Request.Form("pmg_other_pay2"), 0))
pmg_other_pay3 = CLng(f_toString(Request.Form("pmg_other_pay3"), 0))
pmg_tax_yes = CLng(f_toString(Request.Form("pmg_tax_yes"), 0))
pmg_tax_no = CLng(f_toString(Request.Form("pmg_tax_no"), 0))
pmg_tax_reduced = CLng(f_toString(Request.Form("pmg_tax_reduced"), 0))
pmg_give_total = CLng(f_toString(Request.Form("pmg_give_tot"), 0))

'공제항목
de_nps_amt = CLng(f_toString(Request.Form("de_nps_amt"), 0))
de_nhis_amt = CLng(f_toString(Request.Form("de_nhis_amt"), 0))
de_epi_amt = CLng(f_toString(Request.Form("de_epi_amt"), 0))
de_longcare_amt = CLng(f_toString(Request.Form("de_longcare_amt"), 0))
de_income_tax = CLng(f_toString(Request.Form("de_income_tax"), 0))
de_wetax = CLng(f_toString(Request.Form("de_wetax"), 0))
de_other_amt1 = CLng(f_toString(Request.Form("de_other_amt1"), 0))
'de_saving_amt = CLng(Request.Form("de_saving_amt"))
de_saving_amt = 0
de_other_amt2 = 0
de_other_amt3 = 0
de_sawo_amt = CLng(f_toString(Request.Form("de_sawo_amt"), 0))
'de_johab_amt = CLng(Request.Form("de_johab_amt"))
de_johab_amt = 0
de_special_tax = 0
de_hyubjo_amt = CLng(f_toString(Request.Form("de_hyubjo_amt"), 0))
de_school_amt = CLng(f_toString(Request.Form("de_school_amt"), 0))
de_nhis_bla_amt = CLng(f_toString(Request.Form("de_nhis_bla_amt"), 0))
de_long_bla_amt = CLng(f_toString(Request.Form("de_long_bla_amt"), 0))
de_year_incom_tax = CLng(f_toString(Request.Form("de_year_incom_tax"), 0))
de_year_wetax = CLng(f_toString(Request.Form("de_year_wetax"), 0))
de_year_incom_tax2 = CLng(f_toString(Request.Form("de_year_incom_tax2"), 0))
de_year_wetax2 = CLng(f_toString(Request.Form("de_year_wetax2"), 0))

de_deduct_total = CDbl(f_toString(Request.Form("de_deduct_tot"), 0))

DBConn.BeginTrans

If u_type = "U" then
	objBuilder.Append "UPDATE pay_month_give SET "
	objBuilder.Append "	pmg_in_date='"&pmg_in_date&"',cost_group='"&cost_group&"',cost_center='"&cost_center&"',"
	objBuilder.Append "	pmg_base_pay='"&pmg_base_pay&"',pmg_meals_pay ='"&pmg_meals_pay&"',pmg_postage_pay ='"&pmg_postage_pay&"',"
	objBuilder.Append "	pmg_re_pay='"&pmg_re_pay&"',pmg_overtime_pay='"&pmg_overtime_pay&"',pmg_car_pay='"&pmg_car_pay&"',"
	objBuilder.Append "	pmg_position_pay='"&pmg_position_pay&"',pmg_custom_pay='"&pmg_custom_pay&"',pmg_job_pay='"&pmg_job_pay&"',"
	objBuilder.Append "	pmg_job_support='"&pmg_job_support&"',pmg_jisa_pay='"&pmg_jisa_pay&"',pmg_long_pay='"&pmg_long_pay&"',"
	objBuilder.Append "	pmg_disabled_pay='"&pmg_disabled_pay&"',pmg_family_pay='"&pmg_family_pay&"',pmg_school_pay='"&pmg_school_pay&"',"
	objBuilder.Append "	pmg_qual_pay='"&pmg_qual_pay&"',pmg_tax_yes='"&pmg_tax_yes&"',pmg_tax_no='"&pmg_tax_no&"',"
	objBuilder.Append "	pmg_tax_reduced='"&pmg_tax_reduced&"',pmg_give_total='"&pmg_give_total&"',pmg_bank_name='"&pmg_bank_name&"',"
	objBuilder.Append "	pmg_account_no='"&pmg_account_no&"',pmg_account_holder='"&pmg_account_holder&"',pmg_mod_user='"&user_name&"',pmg_mod_date=NOW() "
	objBuilder.Append "WHERE pmg_yymm = '"&pmg_yymm&"' AND pmg_id = '1' AND pmg_emp_no = '"&emp_no&"' AND pmg_company = '"&pmg_company&"';"

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	objBuilder.Append "UPDATE pay_month_deduct SET "
	objBuilder.Append "	cost_group='"&cost_group&"',cost_center='"&cost_center&"',de_nps_amt='"&de_nps_amt&"',"
	objBuilder.Append "	de_nhis_amt ='"&de_nhis_amt&"',de_epi_amt ='"&de_epi_amt&"',de_longcare_amt ='"&de_longcare_amt&"',"
	objBuilder.Append "	de_income_tax='"&de_income_tax&"',de_wetax='"&de_wetax&"',de_other_amt1='"&de_other_amt1&"',"
	objBuilder.Append "	de_saving_amt='"&de_saving_amt&"',de_sawo_amt='"&de_sawo_amt&"',de_johab_amt='"&de_johab_amt&"',"
	objBuilder.Append "	de_hyubjo_amt='"&de_hyubjo_amt&"',de_school_amt='"&de_school_amt&"',de_nhis_bla_amt='"&de_nhis_bla_amt&"',"
	objBuilder.Append "	de_long_bla_amt='"&de_long_bla_amt&"',de_year_incom_tax='"&de_year_incom_tax&"',de_year_wetax='"&de_year_wetax&"',"
	objBuilder.Append "	de_year_incom_tax2='"&de_year_incom_tax2&"',de_year_wetax2='"&de_year_wetax2&"',de_deduct_total='"&de_deduct_total&"',"
	objBuilder.Append "	de_mod_user='"&user_name&"',de_mod_date=NOW() "
	objBuilder.Append "WHERE de_yymm = '"&pmg_yymm&"' AND de_id = '1' AND de_emp_no = '"&emp_no&"' AND de_company = '"&pmg_company&"';"


	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
Else
	objBuilder.Append "INSERT INTO pay_month_give(pmg_yymm,pmg_id,pmg_emp_no,pmg_company,pmg_date,"
	objBuilder.Append "pmg_in_date,pmg_emp_name,pmg_emp_type,pmg_org_code,pmg_org_name,"
	objBuilder.Append "pmg_bonbu,pmg_saupbu,pmg_team,pmg_reside_place,pmg_reside_company,"
	objBuilder.Append "pmg_grade,pmg_position,pmg_base_pay,pmg_meals_pay,pmg_postage_pay,"
	objBuilder.Append "pmg_re_pay,pmg_overtime_pay,pmg_car_pay,pmg_position_pay,pmg_custom_pay,"
	objBuilder.Append "pmg_job_pay,pmg_job_support,pmg_jisa_pay,pmg_long_pay,pmg_disabled_pay,"
	objBuilder.Append "pmg_family_pay,pmg_school_pay,pmg_qual_pay,pmg_other_pay1,pmg_other_pay2,"
	objBuilder.Append "pmg_other_pay3,pmg_tax_yes,pmg_tax_no,pmg_tax_reduced,pmg_give_total,"
	objBuilder.Append "pmg_bank_name,pmg_account_no,pmg_account_holder,cost_group,cost_center,pmg_reg_date,pmg_reg_user)"
	objBuilder.Append "VALUES('"&pmg_yymm&"','1','"&emp_no&"','"&pmg_company&"','"&pmg_date&"',"
	objBuilder.Append "'"&pmg_in_date&"','"&pmg_emp_name&"','"&pmg_emp_type&"','"&pmg_org_code&"','"&pmg_org_name&"',"
	objBuilder.Append "'"&pmg_bonbu&"','"&pmg_saupbu&"','"&pmg_team&"','"&pmg_reside_place&"','"&pmg_reside_company&"',"
	objBuilder.Append "'"&pmg_grade&"','"&pmg_position&"','"&pmg_base_pay&"','"&pmg_meals_pay&"','"&pmg_postage_pay&"',"
	objBuilder.Append "'"&pmg_re_pay&"','"&pmg_overtime_pay&"','"&pmg_car_pay&"','"&pmg_position_pay&"','"&pmg_custom_pay&"',"
	objBuilder.Append "'"&pmg_job_pay&"','"&pmg_job_support&"','"&pmg_jisa_pay&"','"&pmg_long_pay&"','"&pmg_disabled_pay&"',"
	objBuilder.Append "'"&pmg_family_pay&"','"&pmg_school_pay&"','"&pmg_qual_pay&"','"&pmg_other_pay1&"','"&pmg_other_pay2&"',"
	objBuilder.Append "'"&pmg_other_pay3&"','"&pmg_tax_yes&"','"&pmg_tax_no&"','"&pmg_tax_reduced&"','"&pmg_give_total&"',"
	objBuilder.Append "'"&pmg_bank_name&"','"&pmg_account_no&"','"&pmg_account_holder&"','"&cost_group&"','"&cost_center&"',NOW(),'"&user_name&"');"


	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	objBuilder.Append "INSERT INTO pay_month_deduct(de_yymm,de_id,de_emp_no,de_company,de_date,"
	objBuilder.Append "de_emp_name,de_emp_type,de_org_code,de_org_name,de_bonbu,"
	objBuilder.Append "de_saupbu,de_team,de_reside_place,de_reside_company,de_grade,"
	objBuilder.Append "de_position,de_nps_amt,de_nhis_amt,de_epi_amt,de_longcare_amt,"
	objBuilder.Append "de_income_tax,de_wetax,de_year_incom_tax,de_year_wetax,de_year_incom_tax2,"
	objBuilder.Append "de_year_wetax2,de_other_amt1,de_saving_amt,de_sawo_amt,de_johab_amt,"
	objBuilder.Append "de_hyubjo_amt,de_school_amt,de_nhis_bla_amt,de_long_bla_amt,de_deduct_total,"
	objBuilder.Append "cost_group,cost_center,de_reg_date,de_reg_user)"
	objBuilder.Append "VALUES('"&pmg_yymm&"','1','"&emp_no&"','"&pmg_company&"','"&pmg_date&"',"
	objBuilder.Append "'"&pmg_emp_name&"','"&pmg_emp_type&"','"&pmg_org_code&"','"&pmg_org_name&"','"&pmg_bonbu&"',"
	objBuilder.Append "'"&pmg_saupbu&"','"&pmg_team&"','"&pmg_reside_place&"','"&pmg_reside_company&"','"&pmg_grade&"',"
	objBuilder.Append "'"&pmg_position&"','"&de_nps_amt&"','"&de_nhis_amt&"','"&de_epi_amt&"','"&de_longcare_amt&"',"
	objBuilder.Append "'"&de_income_tax&"','"&de_wetax&"','"&de_year_incom_tax&"','"&de_year_wetax&"','"&de_year_incom_tax2&"',"
	objBuilder.Append "'"&de_year_wetax2&"','"&de_other_amt1&"','"&de_saving_amt&"','"&de_sawo_amt&"','"&de_johab_amt&"',"
	objBuilder.Append "'"&de_hyubjo_amt&"','"&de_school_amt&"','"&de_nhis_bla_amt&"','"&de_long_bla_amt&"','"&de_deduct_total&"',"
	objBuilder.Append "'"&cost_group&"','"&cost_center&"',NOW(),'"&user_name&"');"

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
End If

If Err.number <> 0 Then
	DBConn.RollbackTrans
	'end_msg = sms_msg + "저장중 Error가 발생하였습니다...."
	end_msg = "저장 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	'end_msg = sms_msg + "저장되었습니다...."
	end_msg = "정상적으로 처리되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	parent.opener.location.reload();"
Response.Write "	self.close() ;"
Response.Write "</script>"
Response.End
%>
