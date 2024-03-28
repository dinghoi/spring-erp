<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'On Error Resume Next
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
Dim pmg_yymm, view_condi, pmg_yymm_to, pmg_date, rsInsEmp, rsInsHap
Dim epi_emp, epi_com, long_hap, rsPay, st_in_date, rever_year
Dim emp_name, cost_group, cost_center, pmg_in_date, pmg_emp_type, pmg_grade, pmg_position
Dim pmg_company, pmg_bonbu, pmg_saupbu, pmg_team, pmg_org_code, pmg_org_name, pmg_reside_place
Dim pmg_reside_company, pmg_base_pay, pmg_meals_pay, pmg_postage_pay, pmg_re_pay
Dim pmg_overtime_pay, pmg_car_pay, pmg_position_pay, pmg_custom_pay, pmg_job_pay, pmg_job_support, pmg_jisa_pay, pmg_long_pay
Dim pmg_disabled_pay, pmg_family_pay, pmg_school_pay, pmg_qual_pay, pmg_other_pay1, pmg_other_pay2, pmg_other_pay3, pmg_tax_yes
Dim pmg_tax_no, pmg_tax_reduced, pmg_give_total, meals_pay, car_pay, meals_tax_pay, car_tax_pay

Dim de_nps_amt, de_nhis_amt, de_epi_amt, de_longcare_amt
Dim de_income_tax, de_wetax, de_year_incom_tax, de_year_wetax
Dim de_year_incom_tax2, de_year_wetax2, de_other_amt1, de_special_tax
Dim de_saving_amt, de_sawo_amt, de_johab_amt, de_hyubjo_amt
Dim de_school_amt, de_nhis_bla_amt, de_long_bla_amt, de_deduct_total

Dim incom_emp_no, incom_base_pay, incom_meals_pay, incom_overtime_pay
Dim incom_month_amount, incom_family_cnt, incom_nps_amount, incom_nhis_amount
Dim incom_nps, incom_nhis, incom_wife_yn, incom_age20, incom_age60, incom_old, incom_go_yn, incom_long_yn

Dim rs_sod, long_amt, we_tax, epi_amt

Dim rs_bnk, pmg_bank_name, pmg_account_no, pmg_account_holder, end_msg

pmg_yymm = Request.Form("pmg_yymm1")
view_condi = Request.Form("view_condi1")
pmg_yymm_to = Request.Form("pmg_yymm_to1")
pmg_date = Request.Form("to_date1")

'당월 입사/퇴사일이 15일 이전이면 당월 급여대상임
'st_es_date = mid(cstr(pmg_yymm_to),1,4) + "-" + mid(cstr(pmg_yymm_to),5,2) + "-" + "01"
st_in_date = Mid(CStr(pmg_yymm_to), 1, 4)&"-"&Mid(CStr(pmg_yymm_to), 5, 2)&"-"&"16"
rever_year = Mid(CStr(pmg_yymm_to), 1, 4) '귀속년도

'고용보험(실업) 요율
objBuilder.Append "SELECT emp_rate, com_rate FROM pay_insurance "
objBuilder.Append  "WHERE insu_yyyy = '"&rever_year&"' AND insu_id = '5503' AND insu_class = '01';"

Set rsInsEmp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsInsEmp.EOF Then
	epi_emp = FormatNumber(rsInsEmp("emp_rate"), 3)
	epi_com = FormatNumber(rsInsEmp("com_rate"), 3)
Else
	epi_emp = 0
	epi_com = 0
End If
rsInsEmp.Close() : Set rsInsEmp = Nothing

'장기요양보험 요율
objBuilder.Append "SELECT hap_rate FROM pay_insurance "
objBuilder.Append "WHERE insu_yyyy = '"&rever_year&"' AND insu_id = '5504' AND insu_class = '01';"

Set rsInsHap = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsInsHap.EOF Then
	long_hap = FormatNumber(rsInsHap("hap_rate"), 3)
Else
	long_hap = 0
End if
rsInsHap.Close() : Set rsInsHap = Nothing

' 급여지급월의 15일까지 입사자 당월급여처리를 위한 급여데이타 생성(전월급여지급이 없음)
objBuilder.Append "SELECT emtt.emp_no, emtt.emp_name, emtt.emp_company, emtt.cost_group, emtt.cost_center, "

objBuilder.Append "	pmgt.pmg_emp_no, pmgt.pmg_in_date, pmgt.pmg_emp_type, pmgt.pmg_grade, pmgt.pmg_position, "
objBuilder.Append "	pmgt.pmg_company, pmgt.pmg_bonbu, pmgt.pmg_saupbu, pmgt.pmg_team, "
objBuilder.Append "	pmgt.pmg_org_code, pmgt.pmg_org_name, pmgt.pmg_reside_place, pmgt.pmg_reside_company, "
objBuilder.Append "	pmgt.pmg_base_pay, pmgt.pmg_meals_pay, pmgt.pmg_postage_pay, pmgt.pmg_re_pay, "
objBuilder.Append "	pmgt.pmg_overtime_pay, pmgt.pmg_car_pay, pmgt.pmg_position_pay, pmgt.pmg_custom_pay, "
objBuilder.Append "	pmgt.pmg_job_pay, pmgt.pmg_job_support, pmgt.pmg_jisa_pay, pmgt.pmg_long_pay, "
objBuilder.Append "	pmgt.pmg_disabled_pay, pmgt.pmg_family_pay, pmgt.pmg_school_pay, pmgt.pmg_qual_pay, "
objBuilder.Append "	pmgt.pmg_other_pay1, pmgt.pmg_other_pay2, pmgt.pmg_other_pay3, pmgt.pmg_tax_yes, "
objBuilder.Append "	pmgt.pmg_tax_no, pmgt.pmg_tax_reduced, pmgt.pmg_give_total, "

objBuilder.Append "	pmdt.de_emp_no, pmdt.de_nps_amt, pmdt.de_nhis_amt, pmdt.de_epi_amt, pmdt.de_longcare_amt,"
objBuilder.Append "	pmdt.de_income_tax, pmdt.de_wetax, pmdt.de_year_incom_tax, pmdt.de_year_wetax, "
objBuilder.Append "	pmdt.de_year_incom_tax2, pmdt.de_year_wetax2, pmdt.de_other_amt1, pmdt.de_special_tax, "
objBuilder.Append "	pmdt.de_saving_amt, pmdt.de_sawo_amt, pmdt.de_johab_amt, pmdt.de_hyubjo_amt, "
objBuilder.Append "	pmdt.de_school_amt, pmdt.de_nhis_bla_amt, pmdt.de_long_bla_amt, pmdt.de_deduct_total, "

objBuilder.Append "	pyit.incom_emp_no, pyit.incom_base_pay, pyit.incom_meals_pay, pyit.incom_overtime_pay, "
objBuilder.Append "	pyit.incom_month_amount, pyit.incom_family_cnt, pyit.incom_nps_amount, pyit.incom_nhis_amount, "
objBuilder.Append "	pyit.incom_nps, pyit.incom_nhis, pyit.incom_wife_yn, pyit.incom_age20, pyit.incom_age60, "
objBuilder.Append "	pyit.incom_old,	pyit.incom_go_yn, pyit.incom_long_yn "

objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "LEFT OUTER JOIN pay_month_give AS pmgt ON emtt.emp_no = pmgt.pmg_emp_no "
objBuilder.Append "	AND emtt.emp_company = pmgt.pmg_company "
objBuilder.Append "	AND pmgt.pmg_yymm = '"&pmg_yymm&"' AND pmgt.pmg_id = '1' "
objBuilder.Append "LEFT OUTER JOIN pay_month_deduct AS pmdt ON emtt.emp_no = pmdt.de_emp_no "
objBuilder.Append "	AND emtt.emp_company = pmdt.de_company "
objBuilder.Append "	AND pmdt.de_yymm = '"&pmg_yymm&"' AND pmdt.de_id = '1' "
objBuilder.Append "LEFT OUTER JOIN pay_year_income AS pyit ON emtt.emp_no = pyit.incom_emp_no"
objBuilder.Append "	AND pyit.incom_year = '"&pmg_yymm&"' "
objBuilder.Append "WHERE (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date >= '"&st_in_date&"') "
objBuilder.Append "	AND emtt.emp_in_date < '"&st_in_date&"' AND emtt.emp_pay_id <> '5' AND emtt.emp_no < '900000' "

If view_condi <> "전체" Then
	objBuilder.Append "	AND emtt.emp_company = '"&view_condi&"' "
End If

objBuilder.Append "ORDER BY emtt.emp_no;"

Set rsPay = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsPay.EOF Then
	DBConn.BeginTrans

	Do Until rsPay.EOF
		emp_no = rsPay("emp_no")
		emp_company = rsPay("emp_company")
		emp_name = rsPay("emp_name")
		cost_group = rsPay("cost_group")
		cost_center = rsPay("cost_center")

		'급여 정보 데이터
		If f_toString(rsPay("pmg_emp_no"), "") <> "" Then
			pmg_in_date = rsPay("pmg_in_date")
			pmg_emp_type = rsPay("pmg_emp_type")
			pmg_grade = rsPay("pmg_grade")
			pmg_position = rsPay("pmg_position")
			pmg_company = rsPay("pmg_company")
			pmg_bonbu = rsPay("pmg_bonbu")
			pmg_saupbu = rsPay("pmg_saupbu")
			pmg_team = rsPay("pmg_team")
			pmg_org_code = rsPay("pmg_org_code")
			pmg_org_name = rsPay("pmg_org_name")
			pmg_reside_place = rsPay("pmg_reside_place")
			pmg_reside_company = rsPay("pmg_reside_company")

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
			pmg_give_total = rsPay("pmg_give_total")

			'pmg_give_total = pmg_base_pay + pmg_meals_pay + pmg_postage_pay + pmg_re_pay + pmg_overtime_pay + pmg_car_pay + pmg_position_pay + pmg_custom_pay + pmg_job_pay + pmg_job_support + pmg_jisa_pay + pmg_long_pay + pmg_disabled_pay

			meals_pay = pmg_meals_pay
			car_pay = pmg_car_pay
			meals_tax_pay = 0
			car_tax_pay = 0

			If meals_pay > 100000 Then
				meals_tax_pay = meals_pay - 100000
				meals_pay =  100000
			End If

			If car_pay > 200000 Then
				car_tax_pay = car_pay - 200000
				car_pay =  200000
			End If

			'pmg_tax_yes = pmg_base_pay + pmg_postage_pay + pmg_re_pay + pmg_overtime_pay + pmg_position_pay + pmg_custom_pay + pmg_job_pay + pmg_job_support + pmg_jisa_pay + pmg_long_pay + pmg_disabled_pay + meals_tax_pay + car_tax_pay

			pmg_tax_no = meals_pay + car_pay

			If f_toString(rsPay("de_emp_no"), "") <> "" Then
				de_nps_amt = Int(rsPay("de_nps_amt"))
				de_nhis_amt = Int(rsPay("de_nhis_amt"))
				de_epi_amt = Int(rsPay("de_epi_amt"))
				de_longcare_amt = Int(rsPay("de_longcare_amt"))
				de_income_tax = Int(rsPay("de_income_tax"))
				de_wetax = Int(rsPay("de_wetax"))
				de_year_incom_tax = Int(rsPay("de_year_incom_tax"))
				de_year_wetax = Int(rsPay("de_year_wetax"))
				de_year_incom_tax2 = Int(rsPay("de_year_incom_tax2"))
				de_year_wetax2 = Int(rsPay("de_year_wetax2"))
				de_other_amt1 = Int(rsPay("de_other_amt1"))
				de_special_tax = rsPay("de_special_tax")
				de_saving_amt = rsPay("de_saving_amt")
				de_sawo_amt = Int(rsPay("de_sawo_amt"))
				de_johab_amt = rsPay("de_johab_amt")
				de_hyubjo_amt = Int(rsPay("de_hyubjo_amt"))
				de_school_amt = Int(rsPay("de_school_amt"))
				de_nhis_bla_amt = Int(rsPay("de_nhis_bla_amt"))
				de_long_bla_amt = Int(rsPay("de_long_bla_amt"))
				de_deduct_total = Int(rsPay("de_deduct_total"))
			 Else
				de_nps_amt = 0
				de_nhis_amt = 0
				de_epi_amt = 0
				de_longcare_amt = 0
				de_income_tax = 0
				de_wetax = 0
				de_year_incom_tax = 0
				de_year_wetax = 0
				de_year_incom_tax2 = 0
				de_year_wetax2 = 0
				de_other_amt1 = 0
				de_special_tax = 0
				de_saving_amt = 0
				de_sawo_amt = 0
				de_johab_amt = 0
				de_hyubjo_amt = 0
				de_school_amt = 0
				de_nhis_bla_amt = 0
				de_long_bla_amt = 0
				de_deduct_total = 0
			 End If
		 Else '급여 정보 없는 경우
			 pmg_base_pay = 0
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
			 pmg_give_total = 0

			 de_nps_amt = 0
			 de_nhis_amt = 0
			 de_epi_amt = 0
			 de_longcare_amt = 0
			 de_income_tax = 0
			 de_wetax = 0
			 de_year_incom_tax = 0
			 de_year_wetax = 0
			 de_year_incom_tax2 = 0
			 de_year_wetax2 = 0
			 de_other_amt1 = 0
			 de_special_tax = 0
			 de_saving_amt = 0
			 de_sawo_amt = 0
			 de_johab_amt = 0
			 de_hyubjo_amt = 0
			 de_school_amt = 0
			 de_nhis_bla_amt = 0
			 de_long_bla_amt = 0
			 de_deduct_total = 0

			 '기본급/식대등 가져오기
			 incom_family_cnt = 0
			 'Sql = "SELECT * FROM pay_year_income where incom_emp_no = '"&emp_no&"' and incom_year = '"&rever_year&"'"
			 'Set Rs_year = DbConn.Execute(SQL)
			 'if not Rs_year.eof then
			If f_toString(rsPay("incom_emp_no"), "") <> "" Then
				pmg_base_pay = rsPay("incom_base_pay")
				pmg_meals_pay = rsPay("incom_meals_pay")
				pmg_overtime_pay = rsPay("incom_overtime_pay")

				If f_toString(rsPay("incom_month_amount"), 0) = 0 Then
					incom_month_amount = rsPay("incom_base_pay") + rsPay("incom_overtime_pay")
				Else
					incom_month_amount = rsPay("incom_month_amount")
				End If

				incom_family_cnt = rsPay("incom_family_cnt")
				incom_nps_amount = rsPay("incom_nps_amount")
				incom_nhis_amount = rsPay("incom_nhis_amount")
				incom_nps = rsPay("incom_nps")
				incom_nhis = rsPay("incom_nhis")
				incom_wife_yn = Int(rsPay("incom_wife_yn"))
				incom_age20 = rsPay("incom_age20")
				incom_age60 = rsPay("incom_age60")
				incom_old = rsPay("incom_old")
				incom_go_yn = rsPay("incom_go_yn")
				incom_long_yn = rsPay("incom_long_yn")
			Else
				pmg_base_pay = 0
				pmg_meals_pay = 0
				pmg_overtime_pay = 0

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
			End If

			pmg_tax_yes = pmg_base_pay + pmg_overtime_pay
			pmg_tax_no = pmg_meals_pay
			pmg_give_total = pmg_tax_yes + pmg_tax_no

			'if incom_family_cnt = 0 then
				incom_family_cnt = incom_wife_yn + incom_age20 + incom_age60 + incom_old + 1 '부양가족은 본인포함으로
			'end if

			'근로소득 간이세액 산출
			inc_st_amt = 0
			inc_incom = 0

			'Sql = "SELECT * FROM pay_income_amount where ('"&incom_month_amount&"' BETWEEN inc_from_amt and inc_to_amt) and (inc_yyyy = '"&rever_year&"')"
			objBuilder.Append "SELECT inc_st_amt, inc_incom1, inc_incom2, inc_incom3, inc_incom4, inc_incom5 "
			objBuilder.Append "	inc_incom6, inc_incom7, inc_incom8, inc_incom9, inc_incom10, inc_incom11 "
			objBUilder.Append "FROM pay_income_amount "
			objBuilder.Append "WHERE ('"&incom_month_amount&"' BETWEEN inc_from_amt AND inc_to_amt) "
			objBuilder.Append "	AND inc_yyyy = '"&rever_year&"';"

			Set rs_sod = DBConn.Execute(objBuilder.ToString())

			If Not rs_sod.EOF Then
				inc_st_amt = Int(rs_sod("inc_st_amt"))

				If incom_family_cnt = 1 Then
					inc_incom = rs_sod("inc_incom1")
				End If

				If incom_family_cnt = 2 Then
				   inc_incom = rs_sod("inc_incom2")
				End If

				If incom_family_cnt = 3 Then
				   inc_incom = rs_sod("inc_incom3")
				End If

				If incom_family_cnt = 4 Then
				   inc_incom = rs_sod("inc_incom4")
				End If

				If incom_family_cnt = 5 Then
				   inc_incom = rs_sod("inc_incom5")
				End If

				If incom_family_cnt = 6 Then
				   inc_incom = rs_sod("inc_incom6")
				End If

				If incom_family_cnt = 7 Then
				   inc_incom = rs_sod("inc_incom7")
				End If

				If incom_family_cnt = 8 Then
				   inc_incom = rs_sod("inc_incom8")
				End If

				If incom_family_cnt = 9 Then
				   inc_incom = rs_sod("inc_incom9")
				End If

				If incom_family_cnt = 10 Then
				   inc_incom = rs_sod("inc_incom10")
				End If

				If incom_family_cnt = 11 Then
				   inc_incom = rs_sod("inc_incom11")
				End If
			End If
			rs_sod.Close()

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
				epi_amt = pmg_give_tot * (epi_emp / 100)
				epi_amt = Int(epi_amt)
				de_epi_amt = (Int(epi_amt / 10)) * 10
			Else
				de_epi_amt = 0
			End If

			'지방소득세
			we_tax = inc_incom * (10 / 100)
			we_tax = Int(we_tax)
			de_wetax = (Int(we_tax / 10)) * 10

			de_deduct_total = de_nps_amt + de_nhis_amt + de_epi_amt + de_longcare_amt + de_income_tax + de_wetax
			pmg_curr_pay = pmg_give_total - de_deduct_total
		End If

		objBuilder.Append "SELECT bank_name, account_no, account_holder "
		objBuilder.Append "FROM pay_bank_account WHERE emp_no = '"&emp_no&"';"

		Set rs_bnk = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If Not rs_bnk.EOF Then
			pmg_bank_name = rs_bnk("bank_name")
			pmg_account_no = rs_bnk("account_no")
			pmg_account_holder = rs_bnk("account_holder")
		Else
			pmg_bank_name = ""
			pmg_account_no = ""
			pmg_account_holder = ""
		End If
		rs_bnk.Close()

		objBuilder.Append "INSERT INTO pay_month_give(pmg_yymm,pmg_id,pmg_emp_no,pmg_company,pmg_date,"
		objBuilder.Append "pmg_in_date,pmg_emp_name,pmg_emp_type,pmg_org_code,pmg_org_name,"
		objBuilder.Append "pmg_bonbu,pmg_saupbu,pmg_team,pmg_reside_place,pmg_reside_company,"
		objBuilder.Append "pmg_grade,pmg_position,pmg_base_pay,pmg_meals_pay,pmg_postage_pay,"
		objBuilder.Append "pmg_re_pay,pmg_overtime_pay,pmg_car_pay,pmg_position_pay,pmg_custom_pay,"
		objBuilder.Append "pmg_job_pay,pmg_job_support,pmg_jisa_pay,pmg_long_pay,pmg_disabled_pay,"
		objBuilder.Append "pmg_family_pay,pmg_school_pay,pmg_qual_pay,pmg_other_pay1,pmg_other_pay2,"
		objBuilder.Append "pmg_other_pay3,pmg_tax_yes,pmg_tax_no,pmg_tax_reduced,pmg_give_total,"
		objBuilder.Append "pmg_bank_name,pmg_account_no,pmg_account_holder,cost_group,cost_center,"
		objBuilder.Append "pmg_reg_date,pmg_reg_user)"
		objBuilder.Append "VALUES('"&pmg_yymm_to&"','1','"&emp_no&"','"&pmg_company&"','"&pmg_date&"',"
		objBuilder.Append "'"&pmg_in_date&"','"&emp_name&"','"&pmg_emp_type&"','"&pmg_org_code&"','"&pmg_org_name&"',"
		objBuilder.Append "'"&pmg_bonbu&"','"&pmg_saupbu&"','"&pmg_team&"','"&pmg_reside_place&"','"&pmg_reside_company&"',"
		objBuilder.Append "'"&pmg_grade&"','"&pmg_position&"','"&pmg_base_pay&"','"&pmg_meals_pay&"','"&pmg_postage_pay&"',"
		objBuilder.Append "'"&pmg_re_pay&"','"&pmg_overtime_pay&"','"&pmg_car_pay&"','"&pmg_position_pay&"','"&pmg_custom_pay&"',"
		objBuilder.Append "'"&pmg_job_pay&"','"&pmg_job_support&"','"&pmg_jisa_pay&"','"&pmg_long_pay&"','"&pmg_disabled_pay&"',"
		objBuilder.Append "'"&pmg_family_pay&"','"&pmg_school_pay&"','"&pmg_qual_pay&"','"&pmg_other_pay1&"','"&pmg_other_pay2&"',"
		objBuilder.Append "'"&pmg_other_pay3&"','"&pmg_tax_yes&"','"&pmg_tax_no&"','"&pmg_tax_reduced&"','"&pmg_give_total&"',"
		objBuilder.Append "'"&pmg_bank_name&"','"&pmg_account_no&"','"&pmg_account_holder&"','"&cost_group&"','"&cost_center&"',"
		objBuilder.Append "NOW(),'"&user_id&"');"

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
		objBuilder.Append "VALUES('"&pmg_yymm_to&"','1','"&emp_no&"','"&pmg_company&"','"&pmg_date&"',"
		objBuilder.Append "'"&emp_name&"','"&pmg_emp_type&"','"&pmg_org_code&"','"&pmg_org_name&"','"&pmg_bonbu&"',"
		objBuilder.Append "'"&pmg_saupbu&"','"&pmg_team&"','"&pmg_reside_place&"','"&pmg_reside_company&"','"&pmg_grade&"',"
		objBuilder.Append "'"&pmg_position&"','"&de_nps_amt&"','"&de_nhis_amt&"','"&de_epi_amt&"','"&de_longcare_amt&"',"
		objBuilder.Append "'"&de_income_tax&"','"&de_wetax&"','"&de_year_incom_tax&"','"&de_year_wetax&"','"&de_year_incom_tax2&"',"
		objBuilder.Append "'"&de_year_wetax2&"','"&de_other_amt1&"','"&de_saving_amt&"','"&de_sawo_amt&"','"&de_johab_amt&"',"
		objBuilder.Append "'"&de_hyubjo_amt&"','"&de_school_amt&"','"&de_nhis_bla_amt&"','"&de_long_bla_amt&"','"&de_deduct_total&"',"
		objBuilder.Append "'"&cost_group&"','"&cost_center&"',now(),'"&user_id&"');"

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		rsPay.MoveNext()
    Loop
	Set rs_sod = Nothing
	Set rs_bnk = Nothing

	If Err.Number <> "0" Then
		DBConn.RollbackTrans
		end_msg = "처리 중 오류가 발생했습니다."
	Else
		DBConn.CommitTrans
		end_msg = "전월 급여로 당월급여 기초 데이터가 만들어졌습니다."
	End If

	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('"&end_msg&"');"
	Response.Write "	location.replace('/pay/insa_pay_month_batch.asp');"
	Response.Write "</script>"
	Response.End
Else
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('처리할 내역이 없습니다.');"
	Response.Write "	location.replace('/pay/insa_pay_month_batch.asp');"
	Response.Write "</script>"
	Response.End
End If

DBConn.Close() : Set DBConn = Nothing
%>
