<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
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
Dim u_type, emp_name, pmg_yymm, pmg_date, view_condi
Dim rs_bnk, pmg_emp_no, pmg_emp_name, rever_year
Dim bank_name, account_no, account_holder
Dim rsNps, nps_emp, nps_com, nps_from, nps_to
Dim rsNhis, nhis_emp, nhis_com, nhis_from, nhis_to
Dim rsEpi, epi_emp, epi_com, rsHap, long_hap

Dim rsEmp, emp_first_date, emp_in_date, pmg_emp_type, pmg_grade, pmg_position
Dim pmg_company, pmg_bonbu, pmg_saupbu, pmg_team, pmg_org_code, pmg_org_name
Dim pmg_reside_place, pmg_reside_company, cost_center, cost_group
Dim incom_family_cnt, pmg_tax_yes, pmg_tax_no, pmg_give_tot
Dim incom_base_pay, incom_meals_pay, incom_overtime_pay, incom_month_amount, incom_nps_amount
Dim incom_nps, incom_nhis, incom_go_yn, incom_long_yn, incom_wife_yn, incom_age20, incom_age60
Dim incom_old, incom_disab, incom_nhis_amount, title_line, de_deduct_tot
Dim rs_sod, inc_st_amt, inc_incom, de_income_tax, de_nps_amt, de_nhis_amt, long_amt
Dim de_longcare_amt, epi_amt, de_epi_amt, we_tax, de_wetax, pay_curr_amt

Dim pmg_base_pay, pmg_meals_pay, pmg_postage_pay, pmg_re_pay, pmg_overtime_pay
Dim pmg_car_pay, pmg_position_pay, de_other_amt1, pmg_custom_pay, de_sawo_amt
Dim pmg_job_pay, de_hyubjo_amt, pmg_job_support, de_school_amt, pmg_jisa_pay
Dim de_nhis_bla_amt, pmg_long_pay, de_long_bla_amt, pmg_disabled_pay, de_year_incom_tax
Dim pmg_family_pay, de_year_wetax, de_year_incom_tax2, de_year_wetax2, ajaxout
Dim pmg_tax_reduced, rsPay, pmg_school_pay, pmg_qual_pay, pmg_other_pay1, pmg_other_pay2
Dim pmg_other_pay3, pmg_bank_name, pmg_account_no, pmg_account_holder, meals_taxno_pay
Dim car_taxno_pay, meals_tax_pay, car_tax_pay, meals_pay, car_pay, de_special_tax
Dim de_saving_amt, de_johab_amt

u_type = f_Request("u_type")
emp_no = f_Request("emp_no")
pmg_yymm = f_Request("pmg_yymm")
emp_name = f_Request("emp_name")
pmg_date = f_Request("give_date")
view_condi = f_Request("view_condi")

pmg_emp_no = emp_no
pmg_emp_name = emp_name
emp_company = view_condi
rever_year = Mid(CStr(pmg_yymm), 1, 4) '귀속년도

'급여 은행 정보 조회
'Sql = "SELECT * FROM pay_bank_account where emp_no = '"&emp_no&"'"
objBuilder.Append "SELECT bank_name, account_no, account_holder "
objBuilder.Append "FROM pay_bank_account "
objBuilder.Append "WHERE emp_no = '"&emp_no&"';"

Set rs_bnk = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rs_bnk.EOF Then
	bank_name = rs_bnk("bank_name")
	account_no = rs_bnk("account_no")
	account_holder = rs_bnk("account_holder")
Else
	bank_name = ""
	account_no = ""
	account_holder = ""
End If
rs_bnk.Close() : Set rs_bnk = Nothing

'국민연금 요율
'Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&rever_year&"' and insu_id = '5501' and insu_class = '01'"
objBuilder.Append "SELECT emp_rate, com_rate, from_amt, to_amt "
objBuilder.Append "FROM pay_insurance "
objBuilder.Append "WHERE insu_yyyy = '"&rever_year&"' AND insu_id = '5501' AND insu_class = '01';"

Set rsNps = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsNps.EOF Then
	nps_emp = FormatNumber(rsNps("emp_rate"), 3)
	nps_com = FormatNumber(rsNps("com_rate"), 3)
	nps_from = rsNps("from_amt")
	nps_to = rsNps("to_amt")
Else
	nps_emp = 0
	nps_com = 0
	nps_from = 0
	nps_to = 0
End If
rsNps.Close() : Set rsNps = Nothing

'건강보험 요율
'Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&rever_year&"' and insu_id = '5502' and insu_class = '01'"
objBuilder.Append "SELECT emp_rate, com_rate, from_amt, to_amt "
objBuilder.Append "FROM pay_insurance "
objBuilder.Append "WHERE insu_yyyy = '"&rever_year&"' AND insu_id = '5502' AND insu_class = '01';"

Set rsNhis = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsNhis.EOF Then
	nhis_emp = FormatNumber(rsNhis("emp_rate"), 3)
	nhis_com = FormatNumber(rsNhis("com_rate"), 3)
	nhis_from = rsNhis("from_amt")
	nhis_to = rsNhis("to_amt")
Else
	nhis_emp = 0
	nhis_com = 0
	nhis_from = 0
	nhis_to = 0
End If
rsNhis.Close() : Set rsNhis = Nothing

'고용보험(실업) 요율
'Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&rever_year&"' and insu_id = '5503' and insu_class = '01'"
objBuilder.Append "SELECT emp_rate, com_rate "
objBuilder.Append "FROM pay_insurance "
objBuilder.Append "WHERE insu_yyyy = '"&rever_year&"' AND insu_id = '5503' AND insu_class = '01';"

Set rsEpi = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsEpi.EOF Then
	epi_emp = FormatNumber(rsEpi("emp_rate"), 3)
	epi_com = FormatNumber(rsEpi("com_rate"), 3)
Else
	epi_emp = 0
	epi_com = 0
End If
rsEpi.Close() : Set rsEpi = Nothing

'장기요양보험 요율
'Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&rever_year&"' and insu_id = '5504' and insu_class = '01'"
objBuilder.Append "SELECT hap_rate "
objBuilder.Append "FROM pay_insurance "
objBuilder.Append "WHERE insu_yyyy = '"&rever_year&"' AND insu_id = '5504' AND insu_class = '01';"

Set rsHap = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsHap.eof Then
	long_hap = FormatNumber(rsHap("hap_rate"), 3)
Else
	long_hap = 0
End If
rsHap.Close() : Set rsHap = Nothing

'기본급/식대 가져오기
incom_family_cnt = 0

If f_toString(u_type, "") = "" Then
	'Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
	objBuilder.Append "SELECT emtt.emp_first_date, emtt.emp_in_date, emtt.emp_type, emtt.emp_grade, emtt.emp_position, "
	objBuilder.Append "	emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team, emtt.emp_org_code, "
	objBuilder.Append "	emtt.emp_org_name, emtt.emp_reside_place, emtt.emp_reside_company, emtt.cost_center, "
	objBuilder.Append "	emtt.cost_group, "
	objBuilder.Append "	pyit.incom_base_pay, pyit.incom_meals_pay, pyit.incom_overtime_pay, pyit.incom_month_amount, "
	objBuilder.Append "	pyit.incom_family_cnt, pyit.incom_nps_amount, pyit.incom_nhis_amount, pyit.incom_nps, "
	objBuilder.Append "	pyit.incom_nhis, pyit.incom_wife_yn, pyit.incom_age20, pyit.incom_age60, pyit.incom_old, "
	objBuilder.Append "	pyit.incom_disab, pyit.incom_go_yn, pyit.incom_long_yn "
	objBuilder.Append "FROM emp_master AS emtt "
	objBuilder.Append "LEFT OUTER JOIN pay_year_income AS pyit ON emtt.emp_no = pyit.incom_emp_no "
	objBuilder.Append "	AND pyit.incom_year = '"&rever_year&"'"
	objBuilder.Append "WHERE emtt.emp_no = '"&emp_no&"';"

	Set rsEmp = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If Not rsEmp.EOF Then
		emp_first_date = rsEmp("emp_first_date")
		emp_in_date = rsEmp("emp_in_date")
		pmg_emp_type = rsEmp("emp_type")
		pmg_grade = rsEmp("emp_grade")
		pmg_position = rsEmp("emp_position")
		pmg_company = rsEmp("emp_company")
		pmg_bonbu = rsEmp("emp_bonbu")
		pmg_saupbu = rsEmp("emp_saupbu")
		pmg_team = rsEmp("emp_team")
		pmg_org_code = rsEmp("emp_org_code")
		pmg_org_name = rsEmp("emp_org_name")
		pmg_reside_place = rsEmp("emp_reside_place")
		pmg_reside_company = rsEmp("emp_reside_company")
		cost_center = rsEmp("cost_center")
		cost_group = rsEmp("cost_group")

		incom_base_pay = CLng(f_toString(rsEmp("incom_base_pay"), 0))
		incom_meals_pay = CLng(f_toString(rsEmp("incom_meals_pay"), 0))
		incom_overtime_pay = CLng(f_toString(rsEmp("incom_overtime_pay"), 0))

		If f_toString(rsEmp("incom_month_amount"), 0) = 0 Then
			incom_month_amount = CLng(f_toString(rsEmp("incom_base_pay"), 0)) + CLng(f_toString(rsEmp("incom_overtime_pay"), 0))
		Else
			incom_month_amount = CLng(f_toString(rsEmp("incom_month_amount"), 0))
		End If

		incom_family_cnt = CLng(f_toString(rsEmp("incom_family_cnt"), 0))
		incom_nps_amount = CLng(f_toString(rsEmp("incom_nps_amount"), 0))
		incom_nhis_amount = CLng(f_toString(rsEmp("incom_nhis_amount"), 0))
		incom_nps = CLng(f_toString(rsEmp("incom_nps"), 0))
		incom_nhis = CLng(f_toString(rsEmp("incom_nhis"), 0))
		incom_wife_yn = CLng(f_toString(rsEmp("incom_wife_yn"), 0))
		incom_age20 = CLng(f_toString(rsEmp("incom_age20"), 0))
		incom_age60 = CLng(f_toString(rsEmp("incom_age60"), 0))
		incom_old = CLng(f_toString(rsEmp("incom_old"), 0))
		incom_disab = CLng(f_toString(rsEmp("incom_disab"), 0))

		incom_go_yn = f_toString(rsEmp("incom_go_yn"), "여")
		incom_long_yn = f_toString(rsEmp("incom_long_yn"), "여")
	Else
		emp_first_date = ""
		emp_in_date = ""
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
		cost_center = ""
		cost_group = ""

		incom_base_pay = 0
		incom_meals_pay = 0
		incom_overtime_pay = 0
		incom_month_amount = 0
		incom_family_cnt = 0
		incom_nps_amount = 0
		incom_nhis_amount = 0
		incom_nps = 0
		incom_nhis = 0
		incom_go_yn = "여"
		incom_long_yn = "여"
		incom_wife_yn = 0
		incom_age20 = 0
		incom_age60 = 0
		incom_old = 0
		incom_disab = 0
	End If
	rsEmp.Close() : Set rsEmp = Nothing

	pmg_tax_yes = incom_base_pay + incom_overtime_pay
	pmg_tax_no = incom_meals_pay
	pmg_give_tot = pmg_tax_yes + pmg_tax_no

'if incom_family_cnt = 0 then
    incom_family_cnt = incom_wife_yn + incom_age20 + incom_age60 + 1 + incom_age20 + incom_disab'본인포함 및 20세이하/장애인은 추가공제
'end if

	title_line = "급여 지급/공제 입력"
Else
'If u_type = "U" Then
	'sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_emp_no = '"+emp_no+"') and (pmg_company = '"+view_condi+"')"
	objBuilder.Append "SELECT pmgt.pmg_yymm, pmgt.pmg_emp_no, pmgt.pmg_company, pmgt.pmg_date, pmgt.pmg_emp_name, "
	objBuilder.Append "	pmgt.pmg_org_code, pmgt.pmg_org_name, pmgt.pmg_emp_type, pmgt.pmg_grade, pmgt.pmg_position, "
	objBuilder.Append "	pmgt.pmg_base_pay, pmgt.pmg_meals_pay, pmgt.pmg_postage_pay, pmgt.pmg_re_pay, "
	objBuilder.Append "	pmgt.pmg_overtime_pay, pmgt.pmg_car_pay, pmgt.pmg_position_pay, pmgt.pmg_custom_pay, "
	objBuilder.Append "	pmgt.pmg_job_pay, pmgt.pmg_job_support, pmgt.pmg_jisa_pay, pmgt.pmg_long_pay, pmgt.pmg_disabled_pay, "
	objBuilder.Append "	pmgt.pmg_family_pay, pmgt.pmg_school_pay, pmgt.pmg_qual_pay, pmgt.pmg_other_pay1, pmgt.pmg_other_pay2, "
	objBuilder.Append "	pmgt.pmg_other_pay3, pmgt.pmg_tax_yes, pmgt.pmg_tax_no, pmgt.pmg_tax_reduced, pmgt.pmg_give_total, "
	objBuilder.Append "	pmgt.pmg_bank_name, pmgt.pmg_account_no, pmgt.pmg_account_holder, "
	objBuilder.Append "	pmdt.de_emp_no, pmdt.de_nps_amt, pmdt.de_nhis_amt, pmdt.de_epi_amt, pmdt.de_longcare_amt, "
	objBuilder.Append "	pmdt.de_income_tax, pmdt.de_wetax, pmdt.de_year_incom_tax, pmdt.de_year_wetax, pmdt.de_year_incom_tax2, "
	objBuilder.Append "	pmdt.de_year_wetax2, pmdt.de_other_amt1, pmdt.de_special_tax, pmdt.de_saving_amt, pmdt.de_sawo_amt, "
	objBuilder.Append "	pmdt.de_johab_amt, pmdt.de_hyubjo_amt, pmdt.de_school_amt, pmdt.de_nhis_bla_amt, pmdt.de_long_bla_amt "
	objBuilder.Append "FROM pay_month_give AS pmgt "
	objBuilder.Append "LEFT OUTER JOIN pay_month_deduct AS pmdt ON pmgt.pmg_emp_no = pmdt.de_emp_no "
	objBuilder.Append "	AND pmgt.pmg_company= pmdt.de_company "
	objBuilder.Append "	AND pmdt.de_id = '1' AND pmdt.de_yymm = '"&pmg_yymm&"' "
	objBuilder.Append "WHERE pmgt.pmg_yymm = '"&pmg_yymm&"' AND pmgt.pmg_id = '1' "
	objBuilder.Append "	AND pmgt.pmg_emp_no = '"&emp_no&"' AND pmgt.pmg_company = '"&view_condi&"';"

	Set rsPay = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

    pmg_yymm = rsPay("pmg_yymm")
	pmg_emp_no = rsPay("pmg_emp_no")
    pmg_company = rsPay("pmg_company")
	pmg_date = rsPay("pmg_date")
	pmg_emp_name = rsPay("pmg_emp_name")
	pmg_org_code = rsPay("pmg_org_code")
	pmg_org_name = rsPay("pmg_org_name")
	pmg_emp_type = rsPay("pmg_emp_type")
	pmg_grade = rsPay("pmg_grade")
	pmg_position = rsPay("pmg_position")

	pmg_base_pay = rsPay("pmg_base_pay")
	pmg_meals_pay = rsPay("pmg_meals_pay")
	pmg_postage_pay = rsPay("pmg_postage_pay")
	pmg_re_pay = rsPay("pmg_re_pay")
	pmg_overtime_pay = rsPay("pmg_overtime_pay")
	pmg_car_pay = rsPay("pmg_car_pay")
	pmg_position_pay = rsPay("pmg_position_pay")
	pmg_custom_pay = rsPay("pmg_custom_pay")
	pmg_job_pay = rsPay("pmg_job_pay")
	pmg_job_support = rsPay("pmg_job_support")
	pmg_jisa_pay = rsPay("pmg_jisa_pay")
	pmg_long_pay = rsPay("pmg_long_pay")
	pmg_disabled_pay = rsPay("pmg_disabled_pay")
	pmg_family_pay = rsPay("pmg_family_pay")
	pmg_school_pay = rsPay("pmg_school_pay")
	pmg_qual_pay = rsPay("pmg_qual_pay")
	pmg_other_pay1 = rsPay("pmg_other_pay1")
	pmg_other_pay2 = rsPay("pmg_other_pay2")
	pmg_other_pay3 = rsPay("pmg_other_pay3")
	pmg_tax_yes = rsPay("pmg_tax_yes")
	pmg_tax_no = rsPay("pmg_tax_no")
	pmg_tax_reduced = rsPay("pmg_tax_reduced")
	pmg_give_tot = rsPay("pmg_give_total")

	pmg_bank_name = rsPay("pmg_bank_name")
	pmg_account_no = rsPay("pmg_account_no")
	pmg_account_holder = rsPay("pmg_account_holder")

	meals_taxno_pay = pmg_meals_pay
	car_taxno_pay = pmg_car_pay
	meals_tax_pay = 0
	car_tax_pay = 0

	If meals_pay > 100000 Then
	     meals_tax_pay = Int(meals_pay - 100000)
	End If

	If meals_pay > 100000 Then
	     meals_taxno_pay = 100000
	End If

	If car_pay > 200000 Then
	     car_tax_pay = Int(car_pay - 200000)
	End If

	if car_pay > 200000 Then
	     car_taxno_pay =  200000
	End If

	pmg_tax_yes = pmg_base_pay + pmg_postage_pay + pmg_re_pay + pmg_overtime_pay + pmg_position_pay
	pmg_tax_yes = pmg_tax_yes + pmg_custom_pay + pmg_job_pay + pmg_job_support + pmg_jisa_pay + pmg_long_pay
	pmg_tax_yes = pmg_tax_yes + pmg_disabled_pay + meals_tax_pay + car_tax_pay

	pmg_tax_no = meals_taxno_pay + car_taxno_pay

	pmg_give_tot = pmg_tax_yes + pmg_tax_no

	If f_toString(pmg_base_pay, 0) = 0 Then
	    pmg_base_pay =   incom_base_pay
		pmg_meals_pay = incom_meals_pay
		pmg_overtime_pay = incom_overtime_pay
	    pmg_give_tot = incom_base_pay + incom_meals_pay + incom_overtime_pay
	End If

   de_nps_amt = rsPay("de_nps_amt")
   de_nhis_amt = rsPay("de_nhis_amt")
   de_epi_amt = rsPay("de_epi_amt")
   de_longcare_amt = rsPay("de_longcare_amt")
   de_income_tax = rsPay("de_income_tax")
   de_wetax = rsPay("de_wetax")
   de_year_incom_tax = rsPay("de_year_incom_tax")
   de_year_wetax = rsPay("de_year_wetax")
   de_year_incom_tax2 = rsPay("de_year_incom_tax2")
   de_year_wetax2 = rsPay("de_year_wetax2")
   de_other_amt1 = rsPay("de_other_amt1")
   de_special_tax = rsPay("de_special_tax")
   de_saving_amt = rsPay("de_saving_amt")
   de_sawo_amt = rsPay("de_sawo_amt")
   de_johab_amt = rsPay("de_johab_amt")
   de_hyubjo_amt = rsPay("de_hyubjo_amt")
   de_school_amt = rsPay("de_school_amt")
   de_nhis_bla_amt = rsPay("de_nhis_bla_amt")
   de_long_bla_amt = rsPay("de_long_bla_amt")

	If f_toString(rsPay("de_emp_no"), "") <> "" Then
		de_deduct_tot = de_nps_amt + de_nhis_amt + de_epi_amt + de_longcare_amt + de_income_tax
		de_deduct_tot = de_deduct_tot + de_wetax + de_year_incom_tax + de_year_wetax + de_year_incom_tax2 + de_year_wetax2
		de_deduct_tot = de_deduct_tot + de_other_amt1 + de_special_tax + de_saving_amt + de_sawo_amt + de_johab_amt
		de_deduct_tot = de_deduct_tot + de_hyubjo_amt + de_school_amt + de_nhis_bla_amt + de_long_bla_amt
	Else
		de_deduct_tot = 0
	End If
	rsPay.Close() : Set rsPay = Nothing

	pay_curr_amt = pmg_give_tot - de_deduct_tot
	de_deduct_tot = 0

	title_line = "급여 지급/공제 수정"
End If



'근로소득 간이세액 산출
inc_st_amt = 0
inc_incom = 0

' 10000000 < pmg_tax_yes < 14000000 -> 10000000 = pmg_tax_yes
' pmg_tax_yes - 10000000 = cha_a * (98 / 100) = tax_b * (35 / 100) = tax_c 를 더해야 함

' 14000000 > pmg_tax_yes
' 10000000에 해당하는 세액에 + 1372000 +
' (pmg_tax_yes - 14000000) = cha_a * (98 / 100) = tax_b * (35 / 100) = tax_c 를 더해야 함

'Sql = "SELECT * FROM pay_income_amount where ('"&pmg_tax_yes&"' >= inc_from_amt and '"&pmg_tax_yes&"' < inc_to_amt) and (inc_yyyy = '"&rever_year&"')"
objBuilder.Append "SELECT inc_st_amt, inc_incom1, inc_incom2, inc_incom3, inc_incom4, "
objBuilder.Append "	inc_incom5, inc_incom6, inc_incom7, inc_incom8, inc_incom9, inc_incom10, inc_incom11 "
objBuilder.Append "FROM pay_income_amount "
objBuilder.Append "WHERE ('"&pmg_tax_yes&"' >= inc_from_amt AND '"&pmg_tax_yes&"' < inc_to_amt) "
objBuilder.Append "	AND inc_yyyy = '"&rever_year&"';"

Set rs_sod = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rs_sod.EOF Then
   	inc_st_amt = Int(rs_sod("inc_st_amt"))

	Select Case incom_family_cnt
		Case 1
			inc_incom = rs_sod("inc_incom1")
		Case 2
			inc_incom = rs_sod("inc_incom2")
		Case 3
			inc_incom = rs_sod("inc_incom3")
		Case 4
			inc_incom = rs_sod("inc_incom4")
		Case 5
			inc_incom = rs_sod("inc_incom5")
		Case 6
			inc_incom = rs_sod("inc_incom6")
		Case 7
			inc_incom = rs_sod("inc_incom7")
		Case 8
			inc_incom = rs_sod("inc_incom8")
		Case 9
			inc_incom = rs_sod("inc_incom9")
		Case 10
			inc_incom = rs_sod("inc_incom10")
		Case 11
			inc_incom = rs_sod("inc_incom11")
	End Select
End If
rs_sod.Close() : Set rs_sod = Nothing

'소득세
de_income_tax = Int(inc_incom)

'국민연금 계산
'nps_amt = incom_nps_amount * (nps_emp / 100)
'nps_amt = int(nps_amt)
'de_nps_amt = (int(nps_amt / 10)) * 10
de_nps_amt = incom_nps

'건강보험 계산
'nhis_amt = incom_nhis_amount * (nhis_emp / 100)
'nhis_amt = int(nhis_amt)
'de_nhis_amt = (int(nhis_amt / 10)) * 10
de_nhis_amt = incom_nhis

'장기요양보험 계산
If incom_long_yn = "여" Then
	long_amt = de_nhis_amt * (long_hap / 100)
	long_amt = Int(long_amt)
	'long_amt = long_amt / 2
	de_longcare_amt = (Int(long_amt / 10)) * 10
Else
	de_longcare_amt = 0
End If

'고용보험 계산 : 비과세 포함한 금액으로 계산
If incom_go_yn = "여" Then
	'epi_amt = inc_st_amt * (epi_emp / 100)
	epi_amt = pmg_tax_yes * (epi_emp / 100)
	epi_amt = int(epi_amt)
	de_epi_amt = (int(epi_amt / 10)) * 10
else
	de_epi_amt = 0
End If

'지방소득세
we_tax = inc_incom * (10 / 100)
we_tax = Int(we_tax)
de_wetax = (Int(we_tax / 10)) * 10

If f_toString(u_type, "") = "" Then
	de_deduct_tot = de_nps_amt + de_nhis_amt + de_epi_amt + de_longcare_amt + de_income_tax + de_wetax
	pay_curr_amt = pmg_give_tot - de_deduct_tot
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사급여 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%=pmg_date%>" );
			});

			/*$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%'=last_check_date%>" );
			});
			$(function() {    $( "#datepicker2" ).datepicker();
												$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker2" ).datepicker("setDate", "<%'=end_date%>" );
			});
			$(function() {    $( "#datepicker3" ).datepicker();
												$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker3" ).datepicker("setDate", "<%'=car_year%>" );
			});
			*/

			function goAction(){
			   window.close() ;
			}

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.emp_no.value == ""){
					alert('사번을 입력하세요');
					frm.emp_no.focus();
					return false;
				}

				if(document.frm.pmg_date.value == ""){
					alert('급여지급일을 입력하세요');
					frm.pmg_date.focus();
					return false;
				}

				if(document.frm.de_deduct_tot.value == 0){
					alert('세금계산을 하십시요');
					frm.de_deduct_tot.focus();
					return false;
				}


				var result = confirm('저장 하시겠습니까?');
				if (result == true){
					return true;
				}
				return false;
			}

			function give_cal(txtObj){
				base_pay = parseInt(document.frm.pmg_base_pay.value.replace(/,/g,""));
				meals_pay = parseInt(document.frm.pmg_meals_pay.value.replace(/,/g,""));
				postage_pay = parseInt(document.frm.pmg_postage_pay.value.replace(/,/g,""));
				re_pay = parseInt(document.frm.pmg_re_pay.value.replace(/,/g,""));
				overtime_pay = parseInt(document.frm.pmg_overtime_pay.value.replace(/,/g,""));
				car_pay = parseInt(document.frm.pmg_car_pay.value.replace(/,/g,""));
				position_pay = parseInt(document.frm.pmg_position_pay.value.replace(/,/g,""));
				custom_pay = parseInt(document.frm.pmg_custom_pay.value.replace(/,/g,""));
				job_pay = parseInt(document.frm.pmg_job_pay.value.replace(/,/g,""));
				job_support = parseInt(document.frm.pmg_job_support.value.replace(/,/g,""));
				jisa_pay = parseInt(document.frm.pmg_jisa_pay.value.replace(/,/g,""));
				long_pay = parseInt(document.frm.pmg_long_pay.value.replace(/,/g,""));
				disabled_pay = parseInt(document.frm.pmg_disabled_pay.value.replace(/,/g,""));

				e_nps = parseFloat((document.frm.nps_emp.value),3);
				e_nhis = parseFloat((document.frm.nhis_emp.value),3);
				e_epi = parseFloat((document.frm.epi_emp.value),3);
				e_long = parseFloat((document.frm.long_hap.value),3);

		        give_tot = base_pay + meals_pay + postage_pay + re_pay + overtime_pay + car_pay + position_pay + custom_pay + job_pay + job_support + jisa_pay + long_pay + disabled_pay;

				meals_taxno_pay = meals_pay;
				car_taxno_pay = car_pay;
				meals_tax_pay = 0;
				car_tax_pay = 0;

				if (meals_pay > 100000) meals_tax_pay = parseInt(meals_pay - 100000);
				if (meals_pay > 100000) meals_taxno_pay =  100000;
				if (car_pay > 200000) car_tax_pay = parseInt(car_pay - 200000);
				if (car_pay > 200000) car_taxno_pay =  200000;

				tax_yes = base_pay + postage_pay + re_pay + overtime_pay + position_pay + custom_pay + job_pay + job_support + jisa_pay + long_pay + disabled_pay + meals_tax_pay + car_tax_pay;

				tax_no = meals_taxno_pay + car_taxno_pay;

		        base_pay = String(base_pay);
				num_len = base_pay.length;
				sil_len = num_len;
				base_pay = String(base_pay);
				if (base_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) base_pay = base_pay.substr(0,num_len -3) + "," + base_pay.substr(num_len -3,3);
				if (sil_len > 6) base_pay = base_pay.substr(0,num_len -6) + "," + base_pay.substr(num_len -6,3) + "," + base_pay.substr(num_len -2,3);
				document.frm.pmg_base_pay.value = base_pay;

				meals_pay = String(meals_pay);
				num_len = meals_pay.length;
				sil_len = num_len;
				meals_pay = String(meals_pay);
				if (meals_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) meals_pay = meals_pay.substr(0,num_len -3) + "," + meals_pay.substr(num_len -3,3);
				if (sil_len > 6) meals_pay = meals_pay.substr(0,num_len -6) + "," + meals_pay.substr(num_len -6,3) + "," + meals_pay.substr(num_len -2,3);
				document.frm.pmg_meals_pay.value = meals_pay;

				postage_pay = String(postage_pay);
				num_len = postage_pay.length;
				sil_len = num_len;
				postage_pay = String(postage_pay);
				if (postage_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) postage_pay = postage_pay.substr(0,num_len -3) + "," + postage_pay.substr(num_len -3,3);
				if (sil_len > 6) postage_pay = postage_pay.substr(0,num_len -6) + "," + postage_pay.substr(num_len -6,3) + "," + postage_pay.substr(num_len -2,3);
				document.frm.pmg_postage_pay.value = postage_pay;

				re_pay = String(re_pay);
				num_len = re_pay.length;
				sil_len = num_len;
				re_pay = String(re_pay);
				if (re_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) re_pay = re_pay.substr(0,num_len -3) + "," + re_pay.substr(num_len -3,3);
				if (sil_len > 6) re_pay = re_pay.substr(0,num_len -6) + "," + re_pay.substr(num_len -6,3) + "," + re_pay.substr(num_len -2,3);
				document.frm.pmg_re_pay.value = re_pay;

				overtime_pay = String(overtime_pay);
				num_len = overtime_pay.length;
				sil_len = num_len;
				overtime_pay = String(overtime_pay);
				if (overtime_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) overtime_pay = overtime_pay.substr(0,num_len -3) + "," + overtime_pay.substr(num_len -3,3);
				if (sil_len > 6) overtime_pay = overtime_pay.substr(0,num_len -6) + "," + overtime_pay.substr(num_len -6,3) + "," + overtime_pay.substr(num_len -2,3);
				document.frm.pmg_overtime_pay.value = overtime_pay;

				car_pay = String(car_pay);
				num_len = car_pay.length;
				sil_len = num_len;
				car_pay = String(car_pay);
				if (car_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) car_pay = car_pay.substr(0,num_len -3) + "," + car_pay.substr(num_len -3,3);
				if (sil_len > 6) car_pay = car_pay.substr(0,num_len -6) + "," + car_pay.substr(num_len -6,3) + "," + car_pay.substr(num_len -2,3);
				document.frm.pmg_car_pay.value = car_pay;

				position_pay = String(position_pay);
				num_len = position_pay.length;
				sil_len = num_len;
				position_pay = String(position_pay);
				if (position_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) position_pay = position_pay.substr(0,num_len -3) + "," + position_pay.substr(num_len -3,3);
				if (sil_len > 6) position_pay = position_pay.substr(0,num_len -6) + "," + position_pay.substr(num_len -6,3) + "," + position_pay.substr(num_len -2,3);
				document.frm.pmg_position_pay.value = position_pay;

				custom_pay = String(custom_pay);
				num_len = custom_pay.length;
				sil_len = num_len;
				custom_pay = String(custom_pay);
				if (custom_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) custom_pay = custom_pay.substr(0,num_len -3) + "," + custom_pay.substr(num_len -3,3);
				if (sil_len > 6) custom_pay = custom_pay.substr(0,num_len -6) + "," + custom_pay.substr(num_len -6,3) + "," + custom_pay.substr(num_len -2,3);
				document.frm.pmg_custom_pay.value = custom_pay;

				job_pay = String(job_pay);
				num_len = job_pay.length;
				sil_len = num_len;
				job_pay = String(job_pay);
				if (job_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) job_pay = job_pay.substr(0,num_len -3) + "," + job_pay.substr(num_len -3,3);
				if (sil_len > 6) job_pay = job_pay.substr(0,num_len -6) + "," + job_pay.substr(num_len -6,3) + "," + job_pay.substr(num_len -2,3);
				document.frm.pmg_job_pay.value = job_pay;

				job_support = String(job_support);
				num_len = job_support.length;
				sil_len = num_len;
				job_support = String(job_support);
				if (job_support.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) job_support = job_support.substr(0,num_len -3) + "," + job_support.substr(num_len -3,3);
				if (sil_len > 6) job_support = job_support.substr(0,num_len -6) + "," + job_support.substr(num_len -6,3) + "," + job_support.substr(num_len -2,3);
				document.frm.pmg_job_support.value = job_support;

				jisa_pay = String(jisa_pay);
				num_len = jisa_pay.length;
				sil_len = num_len;
				jisa_pay = String(jisa_pay);
				if (jisa_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) jisa_pay = jisa_pay.substr(0,num_len -3) + "," + jisa_pay.substr(num_len -3,3);
				if (sil_len > 6) jisa_pay = jisa_pay.substr(0,num_len -6) + "," + jisa_pay.substr(num_len -6,3) + "," + jisa_pay.substr(num_len -2,3);
				document.frm.pmg_jisa_pay.value = jisa_pay;

				long_pay = String(long_pay);
				num_len = long_pay.length;
				sil_len = num_len;
				long_pay = String(long_pay);
				if (long_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) long_pay = long_pay.substr(0,num_len -3) + "," + long_pay.substr(num_len -3,3);
				if (sil_len > 6) long_pay = long_pay.substr(0,num_len -6) + "," + long_pay.substr(num_len -6,3) + "," + long_pay.substr(num_len -2,3);
				document.frm.pmg_long_pay.value = long_pay;

				disabled_pay = String(disabled_pay);
				num_len = disabled_pay.length;
				sil_len = num_len;
				disabled_pay = String(disabled_pay);
				if (disabled_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) disabled_pay = disabled_pay.substr(0,num_len -3) + "," + disabled_pay.substr(num_len -3,3);
				if (sil_len > 6) disabled_pay = disabled_pay.substr(0,num_len -6) + "," + disabled_pay.substr(num_len -6,3) + "," + disabled_pay.substr(num_len -2,3);
				document.frm.pmg_disabled_pay.value = disabled_pay;

				give_tot = String(give_tot);
				num_len = give_tot.length;
				sil_len = num_len;
				give_tot = String(give_tot);
				if (give_tot.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) give_tot = give_tot.substr(0,num_len -3) + "," + give_tot.substr(num_len -3,3);
				if (sil_len > 6) give_tot = give_tot.substr(0,num_len -6) + "," + give_tot.substr(num_len -6,3) + "," + give_tot.substr(num_len -2,3);
				document.frm.pmg_give_tot.value = give_tot;

				tax_yes = String(tax_yes);
				num_len = tax_yes.length;
				sil_len = num_len;
				tax_yes = String(tax_yes);
				if (tax_yes.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) tax_yes = tax_yes.substr(0,num_len -3) + "," + tax_yes.substr(num_len -3,3);
				if (sil_len > 6) tax_yes = tax_yes.substr(0,num_len -6) + "," + tax_yes.substr(num_len -6,3) + "," + tax_yes.substr(num_len -2,3);
				document.frm.pmg_tax_yes.value = tax_yes;

				tax_no = String(tax_no);
				num_len = tax_no.length;
				sil_len = num_len;
				tax_no = String(tax_no);
				if (tax_no.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) tax_no = tax_no.substr(0,num_len -3) + "," + tax_no.substr(num_len -3,3);
				if (sil_len > 6) tax_no = tax_no.substr(0,num_len -6) + "," + tax_no.substr(num_len -6,3) + "," + tax_no.substr(num_len -2,3);
				document.frm.pmg_tax_no.value = tax_no;
			}

			function deduct_cal(txtObj){
				var give_tot = 0;
				var nps_amt = 0;
				var nhis_amt = 0;
				var epi_amt = 0;
				var long_amt = 0;
				var income_tax = 0;
				var wetax = 0;
				var other_amt1 = 0;
				var sawo_amt = 0;
				var hyubjo_amt = 0;
				var school_amt = 0;
				var long_bal_amt = 0;
				var year_incom_tax = 0;
				var year_wetax = 0;
				var year_incom_tax2 = 0;
				var year_wetax2 = 0;

				give_tot = eval(document.frm.pmg_give_tot.value.replace(/,/g,""));
				nps_amt = eval(document.frm.de_nps_amt.value.replace(/,/g,""));
				nhis_amt = eval(document.frm.de_nhis_amt.value.replace(/,/g,""));
				epi_amt = eval(document.frm.de_epi_amt.value.replace(/,/g,""));
				long_amt = eval(document.frm.de_longcare_amt.value.replace(/,/g,""));
				income_tax = eval(document.frm.de_income_tax.value.replace(/,/g,""));
				wetax = eval(document.frm.de_wetax.value.replace(/,/g,""));
				other_amt1 = eval(document.frm.de_other_amt1.value.replace(/,/g,""));
//				other_amt1 = parseInt(document.frm.de_other_amt1.value.replace(/,/g,""));
				sawo_amt = eval(document.frm.de_sawo_amt.value.replace(/,/g,""));
				hyubjo_amt = eval(document.frm.de_hyubjo_amt.value.replace(/,/g,""));
				school_amt = eval(document.frm.de_school_amt.value.replace(/,/g,""));
				nhis_bal_amt = eval(document.frm.de_nhis_bla_amt.value.replace(/,/g,""));
				long_bal_amt = eval(document.frm.de_long_bla_amt.value.replace(/,/g,""));
				year_incom_tax = eval(document.frm.de_year_incom_tax.value.replace(/,/g,""));
				year_wetax = eval(document.frm.de_year_wetax.value.replace(/,/g,""));
				year_incom_tax2 = eval(document.frm.de_year_incom_tax2.value.replace(/,/g,""));
				year_wetax2 = eval(document.frm.de_year_wetax2.value.replace(/,/g,""));

//				alert(give_tot);
//				alert(other_amt1);

				deduct_tot = 0;
				curr_amt = 0;

				deduct_tot = nps_amt + nhis_amt + epi_amt + long_amt + income_tax + wetax + other_amt1 + sawo_amt + hyubjo_amt + school_amt + nhis_bal_amt + long_bal_amt + year_incom_tax + year_wetax + year_incom_tax2 + year_wetax2;

				curr_amt = give_tot - deduct_tot;

				long_amt = String(long_amt);
				num_len = long_amt.length;
				sil_len = num_len;
				long_amt = String(long_amt);
				if (long_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) long_amt = long_amt.substr(0,num_len -3) + "," + long_amt.substr(num_len -3,3);
				if (sil_len > 6) long_amt = long_amt.substr(0,num_len -6) + "," + long_amt.substr(num_len -6,3) + "," + long_amt.substr(num_len -2,3);
				document.frm.de_longcare_amt.value = long_amt;

				income_tax = String(income_tax);
				num_len = income_tax.length;
				sil_len = num_len;
				income_tax = String(income_tax);
				if (income_tax.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) income_tax = income_tax.substr(0,num_len -3) + "," + income_tax.substr(num_len -3,3);
				if (sil_len > 6) income_tax = income_tax.substr(0,num_len -6) + "," + income_tax.substr(num_len -6,3) + "," + income_tax.substr(num_len -2,3);
				document.frm.de_income_tax.value = income_tax;

				wetax = String(wetax);
				num_len = wetax.length;
				sil_len = num_len;
				wetax = String(wetax);
				if (wetax.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) wetax = wetax.substr(0,num_len -3) + "," + wetax.substr(num_len -3,3);
				if (sil_len > 6) wetax = wetax.substr(0,num_len -6) + "," + wetax.substr(num_len -6,3) + "," + wetax.substr(num_len -2,3);
				document.frm.de_wetax.value = wetax;


				other_amt1 = String(other_amt1);
				num_len = other_amt1.length;
				sil_len = num_len;
				other_amt1 = String(other_amt1);
				if (other_amt1.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) other_amt1 = other_amt1.substr(0,num_len -3) + "," + other_amt1.substr(num_len -3,3);
				if (sil_len > 6) other_amt1 = other_amt1.substr(0,num_len -6) + "," + other_amt1.substr(num_len -6,3) + "," + other_amt1.substr(num_len -2,3);
				eval("document.frm.de_other_amt1.value = other_amt1");

				sawo_amt = String(sawo_amt);
				num_len = sawo_amt.length;
				sil_len = num_len;
				sawo_amt = String(sawo_amt);
				if (sawo_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) sawo_amt = sawo_amt.substr(0,num_len -3) + "," + sawo_amt.substr(num_len -3,3);
				if (sil_len > 6) sawo_amt = sawo_amt.substr(0,num_len -6) + "," + sawo_amt.substr(num_len -6,3) + "," + sawo_amt.substr(num_len -2,3);
				document.frm.de_sawo_amt.value = sawo_amt;

				hyubjo_amt = String(hyubjo_amt);
				num_len = hyubjo_amt.length;
				sil_len = num_len;
				hyubjo_amt = String(hyubjo_amt);
				if (hyubjo_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) hyubjo_amt = hyubjo_amt.substr(0,num_len -3) + "," + hyubjo_amt.substr(num_len -3,3);
				if (sil_len > 6) hyubjo_amt = hyubjo_amt.substr(0,num_len -6) + "," + hyubjo_amt.substr(num_len -6,3) + "," + hyubjo_amt.substr(num_len -2,3);
				document.frm.de_hyubjo_amt.value = hyubjo_amt;

				school_amt = String(school_amt);
				num_len = school_amt.length;
				sil_len = num_len;
				school_amt = String(school_amt);
				if (school_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) school_amt = school_amt.substr(0,num_len -3) + "," + school_amt.substr(num_len -3,3);
				if (sil_len > 6) school_amt = school_amt.substr(0,num_len -6) + "," + school_amt.substr(num_len -6,3) + "," + school_amt.substr(num_len -2,3);
				document.frm.de_school_amt.value = school_amt;

				nhis_bal_amt = String(nhis_bal_amt);
				num_len = nhis_bal_amt.length;
				sil_len = num_len;
				nhis_bal_amt = String(nhis_bal_amt);
				if (nhis_bal_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) nhis_bal_amt = nhis_bal_amt.substr(0,num_len -3) + "," + nhis_bal_amt.substr(num_len -3,3);
				if (sil_len > 6) nhis_bal_amt = nhis_bal_amt.substr(0,num_len -6) + "," + nhis_bal_amt.substr(num_len -6,3) + "," + nhis_bal_amt.substr(num_len -2,3);
				document.frm.de_nhis_bla_amt.value = nhis_bal_amt;

				long_bal_amt = String(long_bal_amt);
				num_len = long_bal_amt.length;
				sil_len = num_len;
				long_bal_amt = String(long_bal_amt);
				if (long_bal_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) long_bal_amt = long_bal_amt.substr(0,num_len -3) + "," + long_bal_amt.substr(num_len -3,3);
				if (sil_len > 6) long_bal_amt = long_bal_amt.substr(0,num_len -6) + "," + long_bal_amt.substr(num_len -6,3) + "," + long_bal_amt.substr(num_len -2,3);
				document.frm.de_long_bla_amt.value = long_bal_amt;

				year_incom_tax = String(year_incom_tax);
				num_len = year_incom_tax.length;
				sil_len = num_len;
				year_incom_tax = String(year_incom_tax);
				if (year_incom_tax.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) year_incom_tax = year_incom_tax.substr(0,num_len -3) + "," + year_incom_tax.substr(num_len -3,3);
				if (sil_len > 6) year_incom_tax = year_incom_tax.substr(0,num_len -6) + "," + year_incom_tax.substr(num_len -6,3) + "," + year_incom_tax.substr(num_len -2,3);
				document.frm.de_year_incom_tax.value = year_incom_tax;

				year_wetax = String(year_wetax);
				num_len = year_wetax.length;
				sil_len = num_len;
				year_wetax = String(year_wetax);
				if (year_wetax.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) year_wetax = year_wetax.substr(0,num_len -3) + "," + year_wetax.substr(num_len -3,3);
				if (sil_len > 6) year_wetax = year_wetax.substr(0,num_len -6) + "," + year_wetax.substr(num_len -6,3) + "," + year_wetax.substr(num_len -2,3);
				document.frm.de_year_wetax.value = year_wetax;

				year_incom_tax2 = String(year_incom_tax2);
				num_len = year_incom_tax2.length;
				sil_len = num_len;
				year_incom_tax2 = String(year_incom_tax2);
				if (year_incom_tax2.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) year_incom_tax2 = year_incom_tax2.substr(0,num_len -3) + "," + year_incom_tax2.substr(num_len -3,3);
				if (sil_len > 6) year_incom_tax2 = year_incom_tax2.substr(0,num_len -6) + "," + year_incom_tax2.substr(num_len -6,3) + "," + year_incom_tax2.substr(num_len -2,3);
				document.frm.de_year_incom_tax2.value = year_incom_tax2;

				year_wetax2 = String(year_wetax2);
				num_len = year_wetax2.length;
				sil_len = num_len;
				year_wetax2 = String(year_wetax2);
				if (year_wetax2.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) year_wetax2 = year_wetax2.substr(0,num_len -3) + "," + year_wetax2.substr(num_len -3,3);
				if (sil_len > 6) year_wetax2 = year_wetax2.substr(0,num_len -6) + "," + year_wetax2.substr(num_len -6,3) + "," + year_wetax2.substr(num_len -2,3);
				document.frm.de_year_wetax2.value = year_wetax2;

				deduct_tot = String(deduct_tot);
				num_len = deduct_tot.length;
				sil_len = num_len;
				deduct_tot = String(deduct_tot);
				if (deduct_tot.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) deduct_tot = deduct_tot.substr(0,num_len -3) + "," + deduct_tot.substr(num_len -3,3);
				if (sil_len > 6) deduct_tot = deduct_tot.substr(0,num_len -6) + "," + deduct_tot.substr(num_len -6,3) + "," + deduct_tot.substr(num_len -2,3);
				if (sil_len > 9) deduct_tot = deduct_tot.substr(0,num_len -9) + "," + deduct_tot.substr(num_len -9,3) + "," + deduct_tot.substr(num_len -5,3) + "," + deduct_tot.substr(num_len -1,3);
				eval("document.frm.de_deduct_tot.value = deduct_tot");

				curr_amt = String(curr_amt);
				num_len = curr_amt.length;
				sil_len = num_len;
				curr_amt = String(curr_amt);
				if (curr_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) curr_amt = curr_amt.substr(0,num_len -3) + "," + curr_amt.substr(num_len -3,3);
				if (sil_len > 6) curr_amt = curr_amt.substr(0,num_len -6) + "," + curr_amt.substr(num_len -6,3) + "," + curr_amt.substr(num_len -2,3);
				if (sil_len > 9) curr_amt = curr_amt.substr(0,num_len -9) + "," + curr_amt.substr(num_len -9,3) + "," + curr_amt.substr(num_len -5,3) + "," + curr_amt.substr(num_len -1,3);
				eval("document.frm.pay_curr_amt.value = curr_amt");

			if (txtObj.value.length<1) {
				txtObj.value=txtObj.value.replace(/,/g,"");
				txtObj.value=txtObj.value.replace(/\D/g,"");
			    }
			var num = txtObj.value;
			if (num == "--" ||  num == "." ) num = "";
			if (num != "" ) {
				temp=new String(num);
				if(temp.length<1) return "";

				// 음수처리
				if(temp.substr(0,1)=="-") minus="-";
					else minus="";

				// 소수점이하처리
				dpoint=temp.search(/\./);

				if(dpoint>0)
				{
				// 첫번째 만나는 .을 기준으로 자르고 숫자제외한 문자 삭제
				dpointVa="."+temp.substr(dpoint).replace(/\D/g,"");
				temp=temp.substr(0,dpoint);
				}else dpointVa="";

				// 숫자이외문자 삭제
				temp=temp.replace(/\D/g,"");
				zero=temp.search(/[1-9]/);

				if(zero==-1) return "";
				else if(zero!=0) temp=temp.substr(zero);

				if(temp.length<4) return minus+temp+dpointVa;
				buf="";
				while (true)
				{
				if(temp.length<3) { buf=temp+buf; break; }

				buf=","+temp.substr(temp.length-3)+buf;
				temp=temp.substr(0, temp.length-3);
				}
				if(buf.substr(0,1)==",") buf=buf.substr(1);

				//return minus+buf+dpointVa;
				txtObj.value = minus+buf+dpointVa;
			}else txtObj.value = "0";
		}

		function taxtax_cal() {
			if (frm.pmg_base_pay.value == 0)
			{
				alert("지급액을 입력하세요");
				frm.pmg_base_pay.focus();
				return;
			}

			var dataString = $("form").serialize();

			$.ajax({
				type: "POST"
				//,url : "/pay/insa_pay_tax_cal.asp"
				,url : "/insa_pay_tax_cal.asp"
				,data: dataString //파라메터
				,success: whenSuccess //성공시 callback
				//error: whenError //실패시 callback
				, error: function(request, status, error){
					console.log("code = "+ request.status + " message = " + request.responseText + " error = " + error);
				}
			});
		}

		function whenSuccess(resdata){
			var aa = resdata.split('|');

			console.log(aa);

			$("div#ajaxout").html(aa[0]);
			//frm.test11.value = aa[1];
			frm.de_epi_amt.value = setComma(aa[2]);
			frm.de_income_tax.value = setComma(aa[3]);
			frm.de_wetax.value = setComma(aa[4]);
			frm.de_nps_amt.value = setComma(aa[5]);
			frm.de_nhis_amt.value = setComma(aa[6]);
			frm.de_longcare_amt.value = setComma(aa[7]);

			give_tot = eval(document.frm.pmg_give_tot.value.replace(/,/g,""));
			nps_amt = eval(document.frm.de_nps_amt.value.replace(/,/g,""));
			nhis_amt = eval(document.frm.de_nhis_amt.value.replace(/,/g,""));
			epi_amt = eval(document.frm.de_epi_amt.value.replace(/,/g,""));
			long_amt = eval(document.frm.de_longcare_amt.value.replace(/,/g,""));
			income_tax = eval(document.frm.de_income_tax.value.replace(/,/g,""));
			wetax = eval(document.frm.de_wetax.value.replace(/,/g,""));
			other_amt1 = eval(document.frm.de_other_amt1.value.replace(/,/g,""));
			sawo_amt = eval(document.frm.de_sawo_amt.value.replace(/,/g,""));
			hyubjo_amt = eval(document.frm.de_hyubjo_amt.value.replace(/,/g,""));
			school_amt = eval(document.frm.de_school_amt.value.replace(/,/g,""));
			nhis_bal_amt = eval(document.frm.de_nhis_bla_amt.value.replace(/,/g,""));
			long_bal_amt = eval(document.frm.de_long_bla_amt.value.replace(/,/g,""));
			year_incom_tax = eval(document.frm.de_year_incom_tax.value.replace(/,/g,""));
			year_wetax = eval(document.frm.de_year_wetax.value.replace(/,/g,""));
			year_incom_tax2 = eval(document.frm.de_year_incom_tax2.value.replace(/,/g,""));
			year_wetax2 = eval(document.frm.de_year_wetax2.value.replace(/,/g,""));

			deduct_tot = 0;
			curr_amt = 0;

			deduct_tot = nps_amt + nhis_amt + epi_amt + long_amt + income_tax + wetax + other_amt1 + sawo_amt + hyubjo_amt + school_amt + nhis_bal_amt + long_bal_amt + year_incom_tax + year_wetax + year_incom_tax2 + year_wetax2;

			curr_amt = give_tot - deduct_tot;

			deduct_tot = String(deduct_tot);
			num_len = deduct_tot.length;
			sil_len = num_len;
			deduct_tot = String(deduct_tot);
			if (deduct_tot.substr(0,1) == "-") sil_len = num_len - 1;
			if (sil_len > 3) deduct_tot = deduct_tot.substr(0,num_len -3) + "," + deduct_tot.substr(num_len -3,3);
			if (sil_len > 6) deduct_tot = deduct_tot.substr(0,num_len -6) + "," + deduct_tot.substr(num_len -6,3) + "," + deduct_tot.substr(num_len -2,3);

			document.frm.de_deduct_tot.value = deduct_tot;

			curr_amt = String(curr_amt);
			num_len = curr_amt.length;
			sil_len = num_len;
			curr_amt = String(curr_amt);
			if (curr_amt.substr(0,1) == "-") sil_len = num_len - 1;
			if (sil_len > 3) curr_amt = curr_amt.substr(0,num_len -3) + "," + curr_amt.substr(num_len -3,3);
			if (sil_len > 6) curr_amt = curr_amt.substr(0,num_len -6) + "," + curr_amt.substr(num_len -6,3) + "," + curr_amt.substr(num_len -2,3);
			document.frm.pay_curr_amt.value = curr_amt;
		}

		function whenError(){
			alert("Error");
		}

		function setComma(str){
			str = ""+str+"";
			var retValue = "";

			for(i=0; i<str.length; i++){
				if(i > 0 && (i%3)==0) {
				   retValue = str.charAt(str.length - i -1) + "," + retValue;
				}else{
				   retValue = str.charAt(str.length - i -1) + retValue;
				}
			}
			return retValue;
		}
	 </script>
	 <style type="text/css">
			.no-input{
				color:gray;
				background-color:#E0E0E0;
				border:1px solid #999999;
			}
		</style>
	</head>
	<body>
		<div id="container">
			<h3 class="insa"><%=title_line%></h3><br/>
			<form action="/pay/insa_pay_month_give_save.asp" method="post" name="frm">
				<input type="hidden" name="rever_year" value="<%=rever_year%>"/>
				<input type="hidden" name="incom_family_cnt" value="<%=incom_family_cnt%>"/>
				<input type="hidden" name="ajaxout" id="ajaxout" size="14" value="<%=ajaxout%>"/>

			<div class="gView">
				<table cellpadding="0" cellspacing="0" class="tableWrite">
					<colgroup>
						<col width="20%" >
						<col width="30%" >
						<col width="20%" >
						<col width="*" >
					</colgroup>
					<tbody>
						<tr>
							<th class="first">사번</th>
							<td class="left">
								<input type="text" name="emp_no" value="<%=pmg_emp_no%>" style="width:90px;" class="no-input" readonly/>
							</td>
							<th >성명</th>
							<td class="left" >
								<input type="text" name="pmg_emp_name" value="<%=pmg_emp_name%>" style="width:90px;" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first">직급</th>
							<td class="left">
								<input type="text" name="pmg_grade" value="<%=pmg_grade%>" style="width:90px;" class="no-input" readonly/>
							</td>
							<th >직책</th>
							<td class="left" >
								<input type="text" name="pmg_position" value="<%=pmg_position%>" style="width:90px;" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first">귀속년월</th>
							<td class="left" >
								<input type="text" name="pmg_yymm" value="<%=pmg_yymm%>" style="width:70px;" readonly/>
							</td>
							<th >지급일</th>
							<td class="left">
								<input type="text" name="pmg_date" value="<%=pmg_date%>" style="width:70px;" id="datepicker" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first">소속</th>
							<td class="left"><%=pmg_company%>&nbsp;&nbsp;<%=pmg_org_name%>(<%=pmg_org_code%>)&nbsp;</td>
							<th>계좌번호</th>
							<td class="left"><%=account_no%>(<%=bank_name%>-<%=account_holder%>)&nbsp;</td>
						</tr>
						<tr>
							<th colspan="2" class="first" style="background:#F5FFFA">지급항목</th>
							<th colspan="2" class="first" style="background:#F8F8FF">공제항목</th>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA;">기본급</th>
							<td class="left">
								<input type="text" name="pmg_base_pay" value="<%=FormatNumber(pmg_base_pay, 0)%>" style="width:100px;text-align:right;" onKeyUp="give_cal(this);"/>
							</td>
							<th style="background:#F8F8FF">국민연금</th>
							<td class="left">
								<input type="text" name="de_nps_amt" value="<%=FormatNumber(de_nps_amt, 0)%>" style="width:100px;text-align:right;" onKeyUp="deduct_cal(this);"/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA;">식대</th>
							<td class="left">
								<input type="text" name="pmg_meals_pay" value="<%=FormatNumber(pmg_meals_pay, 0)%>" style="width:100px;text-align:right;" onKeyUp="give_cal(this);"/>
							</td>
							<th style="background:#F8F8FF">건강보험</th>
							<td class="left">
								<input type="text" name="de_nhis_amt" value="<%=FormatNumber(de_nhis_amt, 0)%>" style="width:100px;text-align:right;" onKeyUp="deduct_cal(this);"/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA;">통신비</th>
							<td class="left">
								<input type="text" name="pmg_postage_pay" value="<%=FormatNumber(pmg_postage_pay, 0)%>" style="width:100px;text-align:right;" onKeyUp="give_cal(this);"/>
							</td>
							<th style="background:#F8F8FF">고용보험</th>
							<td class="left">
								<input type="text" name="de_epi_amt" value="<%=FormatNumber(de_epi_amt, 0)%>" style="width:100px;text-align:right;" onKeyUp="deduct_cal(this);"/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA;">소급급여</th>
							<td class="left">
								<input type="text" name="pmg_re_pay" value="<%=FormatNumber(pmg_re_pay, 0)%>" style="width:100px;text-align:right;" onKeyUp="give_cal(this);"/>
							</td>
							<th style="background:#F8F8FF">장기요양보험</th>
							<td class="left">
								<input type="text" name="de_longcare_amt" value="<%=FormatNumber(de_longcare_amt, 0)%>" style="width:100px;text-align:right;" onKeyUp="deduct_cal(this);"/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA;">연장근로수당</th>
							<td class="left">
								<input type="text" name="pmg_overtime_pay" value="<%=FormatNumber(pmg_overtime_pay, 0)%>" style="width:100px;text-align:right;" onKeyUp="give_cal(this);"/>
							</td>
							<th style="background:#F8F8FF">소득세</th>
							<td class="left">
								<input type="text" name="de_income_tax" value="<%=FormatNumber(de_income_tax, 0)%>" style="width:100px;text-align:right;" onKeyUp="deduct_cal(this);"/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA;">주차지원금</th>
							<td class="left">
								<input type="text" name="pmg_car_pay" value="<%=FormatNumber(pmg_car_pay, 0)%>" style="width:100px;text-align:right;" onKeyUp="give_cal(this);"/>
							</td>
							<th style="background:#F8F8FF">지방소득세</th>
							<td class="left">
								<input type="text" name="de_wetax" value="<%=FormatNumber(de_wetax, 0)%>" style="width:100px;text-align:right;" onKeyUp="deduct_cal(this);"/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA;">직책수당</th>
							<td class="left">
								<input type="text" name="pmg_position_pay" value="<%=FormatNumber(pmg_position_pay, 0)%>" style="width:100px;text-align:right;" onKeyUp="give_cal(this);"/>
							</td>
							<th style="background:#F8F8FF">기타공제</th>
							<td class="left">
								<input type="text" name="de_other_amt1" value="<%=FormatNumber(de_other_amt1, 0)%>" style="width:100px;text-align:right;" onKeyUp="deduct_cal(this);"/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA;">고객관리수당</th>
							<td class="left">
								<input type="text" name="pmg_custom_pay" value="<%=FormatNumber(pmg_custom_pay, 0)%>" style="width:100px;text-align:right;" onKeyUp="give_cal(this);"/>
							</td>
							<th style="background:#F8F8FF">경조회비</th>
							<td class="left">
								<input type="text" name="de_sawo_amt" value="<%=FormatNumber(de_sawo_amt, 0)%>" style="width:100px;text-align:right;" onKeyUp="deduct_cal(this);"/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA;">직무보조비</th>
							<td class="left">
								<input type="text" name="pmg_job_pay" value="<%=FormatNumber(pmg_job_pay, 0)%>" style="width:100px;text-align:right;" onKeyUp="give_cal(this);"/>
							</td>
							<th style="background:#F8F8FF">협조비</th>
							<td class="left">
								<input type="text" name="de_hyubjo_amt" value="<%=FormatNumber(de_hyubjo_amt, 0)%>" style="width:100px;text-align:right;" onKeyUp="deduct_cal(this);"/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA;">업무장려비</th>
							<td class="left">
								<input type="text" name="pmg_job_support" value="<%=FormatNumber(pmg_job_support, 0)%>" style="width:100px;text-align:right;" onKeyUp="give_cal(this);"/>
							</td>
							<th style="background:#F8F8FF;">학자금대출</th>
							<td class="left">
								<input type="text" name="de_school_amt" value="<%=FormatNumber(de_school_amt, 0)%>" style="width:100px;text-align:right;" onKeyUp="deduct_cal(this);"/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA;">본지사근무비</th>
							<td class="left">
								<input type="text" name="pmg_jisa_pay" value="<%=FormatNumber(pmg_jisa_pay, 0)%>" style="width:100px;text-align:right;" onKeyUp="give_cal(this);"/>
							</td>
							<th style="background:#F8F8FF">건강보험료정산</th>
							<td class="left">
								<input type="text" name="de_nhis_bla_amt" value="<%=FormatNumber(de_nhis_bla_amt, 0)%>" style="width:100px;text-align:right;" onKeyUp="deduct_cal(this);"/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA;">근속수당</th>
							<td class="left">
								<input type="text" name="pmg_long_pay" value="<%=FormatNumber(pmg_long_pay, 0)%>" style="width:100px;text-align:right;" onKeyUp="give_cal(this);"/>
							</td>
							<th style="background:#F8F8FF">장기요양보험정산</th>
							<td class="left">
								<input type="text" name="de_long_bla_amt" value="<%=FormatNumber(de_long_bla_amt, 0)%>" style="width:100px;text-align:right;" onKeyUp="deduct_cal(this);"/>
							</td>
						</tr>
						<tr>
							<th class="first" style=" border-bottom:1px solid #e3e3e3; background:#F5FFFA;">장애인수당</th>
							<td class="left">
								<input type="text" name="pmg_disabled_pay" value="<%=FormatNumber(pmg_disabled_pay, 0)%>" style="width:100px;text-align:right;" onKeyUp="give_cal(this);"/>
							</td>
							<th style="background:#F8F8FF">연말정산소득세</th>
							<td class="left">
								<input type="text" name="de_year_incom_tax" value="<%=FormatNumber(de_year_incom_tax, 0)%>" style="width:100px;text-align:right;" onKeyUp="deduct_cal(this);"/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA;"></th>
							<td class="left">&nbsp;
								<input type="hidden" name="pmg_family_pay" value="<%=FormatNumber(pmg_family_pay, 0)%>" style="width:100px;text-align:right;" onKeyUp="give_cal(this);"/>
							</td>
							<th style="background:#F8F8FF">연말정산지방세</th>
							<td class="left">
								<input type="text" name="de_year_wetax" value="<%=FormatNumber(de_year_wetax, 0)%>" style="width:100px;text-align:right;" onKeyUp="deduct_cal(this);"/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA;">과세</th>
							<td class="left">
								<input type="text" name="pmg_tax_yes" value="<%=FormatNumber(pmg_tax_yes, 0)%>" style="width:100px;text-align:right;" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">연말재정산소득세</th>
							<td class="left">
								<input type="text" name="de_year_incom_tax2" value="<%=FormatNumber(de_year_incom_tax2, 0)%>" style="width:100px;text-align:right;" onKeyUp="deduct_cal(this);"/>
								<!--<input type="hidden" name="test11" value="<%'=test11%>" style="width:100px;text-align:center;">-->
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA;">비과세</th>
							<td class="left">
								<input type="text" name="pmg_tax_no" value="<%=FormatNumber(pmg_tax_no,0)%>" style="width:100px;text-align:right;" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">연말재정산지방세</th>
							<td class="left">
								<input type="text" name="de_year_wetax2" value="<%=FormatNumber(de_year_wetax2, 0)%>" style="width:100px;text-align:right;" onKeyUp="deduct_cal(this);"/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA;">감면소득</th>
							<td class="left">
								<input type="text" name="pmg_tax_reduced" value="<%=FormatNumber(pmg_tax_reduced, 0)%>" style="width:100px;text-align:right;" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF;">공제액 계</th>
							<td class="left">
								<input type="text" name="de_deduct_tot" value="<%=FormatNumber(de_deduct_tot, 0)%>" style="width:100px;text-align:right;" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA;">지급액 계</th>
							<td class="left">
								<input type="text" name="pmg_give_tot" value="<%=FormatNumber(pmg_give_tot, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">차인지급액</th>
							<td class="left">
								<input type="text" name="pay_curr_amt" value="<%=FormatNumber(pay_curr_amt, 0)%>" style="width:100px;text-align:right;" class="no-input" readonly/>
								<a href="#" onClick="javascript:taxtax_cal();" class="btn-gray2">세금계산</a>
							</td>
						</tr>
				  </tbody>
				</table>
			</div>
			<br>
			<div align="center">
				<span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();"/></span>
				<span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"/></span>
			</div>
			<input type="hidden" name="u_type" value="<%=u_type%>"/>
			<input type="hidden" name="emp_in_date" value="<%=emp_in_date%>"/>
			<input type="hidden" name="pmg_company" value="<%=pmg_company%>"/>
			<input type="hidden" name="pmg_bonbu" value="<%=pmg_bonbu%>"/>
			<input type="hidden" name="pmg_saupbu" value="<%=pmg_saupbu%>"/>
			<input type="hidden" name="pmg_team" value="<%=pmg_team%>"/>
			<input type="hidden" name="pmg_reside_place" value="<%=pmg_reside_place%>"/>
			<input type="hidden" name="pmg_reside_company" value="<%=pmg_reside_company%>"/>
			<input type="hidden" name="cost_group" value="<%=cost_group%>"/>
			<input type="hidden" name="cost_center" value="<%=cost_center%>"/>
			<input type="hidden" name="pmg_org_name" value="<%=pmg_org_name%>"/>
			<input type="hidden" name="pmg_org_code" value="<%=pmg_org_code%>"/>
			<input type="hidden" name="pmg_emp_type" value="<%=pmg_emp_type%>"/>
			<input type="hidden" name="pmg_bank_name" value="<%=bank_name%>"/>
			<input type="hidden" name="pmg_account_no" value="<%=account_no%>"/>
			<input type="hidden" name="pmg_account_holder" value="<%=account_holder%>"/>
			<input type="hidden" name="nps_emp" value="<%=formatnumber(nps_emp,3)%>"/>
			<input type="hidden" name="nps_com" value="<%=formatnumber(nps_com,3)%>"/>
			<input type="hidden" name="nhis_emp" value="<%=formatnumber(nhis_emp,3)%>"/>
			<input type="hidden" name="nhis_com" value="<%=formatnumber(nhis_com,3)%>"v>
			<input type="hidden" name="epi_emp" value="<%=formatnumber(epi_emp,3)%>"/>
			<input type="hidden" name="epi_com" value="<%=formatnumber(epi_com,3)%>"/>
			<input type="hidden" name="long_hap" value="<%=formatnumber(long_hap,3)%>"/>
			<input type="hidden" name="nps_from" value="<%=nps_from%>"/>
			<input type="hidden" name="nps_to" value="<%=nps_to%>"/>
			<input type="hidden" name="nhis_from" value="<%=nhis_from%>"/>
			<input type="hidden" name="nhis_to" value="<%=nhis_to%>"/>
			<input type="hidden" name="inc_st_amt" value="<%=inc_st_amt%>"/>
			<input type="hidden" name="inc_incom" value="<%=inc_incom%>"/>
			</form>
		</div>
	</body>
</html>