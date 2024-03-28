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
Dim pmg_company, pmg_yymm, objFile, w_cnt, pmg_date
Dim rowcount, xgr, fldcount, tot_cnt, i, j, cn, rs
Dim dz_id, rs_emp, emp_name, emp_bonbu, emp_saupbu, emp_team
Dim emp_org_code, emp_org_name, emp_reside_place, emp_reside_company
Dim emp_in_date, emp_grade, emp_position, emp_type, cost_center, cost_group

Dim pmg_base_pay, pmg_meals_pay, pmg_postage_pay, pmg_re_pay, pmg_overtime_pay
Dim pmg_car_pay, pmg_job_pay, pmg_job_support, pmg_jisa_pay, pmg_disabled_pay
Dim pmg_research_pay, pmg_position_pay, pmg_long_pay, de_nps_amt, de_nhis_amt
Dim de_epi_amt, de_longcare_amt, de_income_tax, de_wetax, de_year_incom_tax
Dim de_year_wetax, de_other_amt1, de_sawo_amt, de_school_amt, de_nhis_bla_amt
Dim de_long_bla_amt, de_hyubjo_amt, de_year_incom_tax2, de_year_wetax2, pmg_family_pay
Dim pmg_school_pay, pmg_qual_pay, pmg_other_pay1, pmg_other_pay2, pmg_other_pay3
Dim pmg_tax_yes, pmg_tax_no, pmg_tax_reduced, pmg_give_total, pmg_custom_pay, meals_pay
Dim car_pay, meals_tax_pay, car_tax_pay, de_special_tax, de_saving_amt, de_johab_amt
Dim de_deduct_total, bank_name, account_no, account_holder

Dim rs_give, rs_dct, end_msg

objFile = Request.Form("objFile")
pmg_company = Request.Form("pmg_company")
pmg_yymm = Request.Form("pmg_yymm")
pmg_date = Request.Form("pmg_date")

w_cnt = 0

Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

DBConn.BeginTrans

cn.Open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
rs.Open "select * from [2:10000]",cn,"0"

rowcount = -1
xgr = rs.getRows
rowcount = UBound(xgr,2)
fldcount = rs.fields.count

tot_cnt = rowcount + 1

If rowcount > -1 Then
	For i=0 To rowcount
		If f_toString(xgr(0,i), "") = "" Then
			Exit For
		End If

		'pmg_company = xgr(7,i)
		'pmg_yymm = xgr(1,i)'귀속년월
		'pmg_date = xgr(2,i)'지급일

		'사번체크
		dz_id = xgr(0, i)

		objBuilder.Append "SELECT emtt.emp_no, emtt.emp_name, emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team, "
		objBuilder.Append "	emtt.emp_org_code, emtt.emp_org_name, emtt.emp_reside_place, emtt.emp_reside_company, "
		objBuilder.Append "	emtt.emp_in_date, emtt.emp_grade, emtt.emp_position, emtt.emp_type, emtt.cost_center, emtt.cost_group, "

		objBuilder.Append "	pbat.bank_name, pbat.account_no, pbat.account_holder "
		objBuilder.Append "FROM emp_master AS emtt "
		objBuilder.Append "INNER JOIN dz_pay_info AS dpit ON emtt.emp_no = dpit.emp_no "
		objBuilder.Append "LEFT OUTER JOIN pay_bank_account AS pbat ON emtt.emp_no = pbat.emp_no "
		objBuilder.Append "WHERE dpit.dz_id='"&dz_id&"' AND dpit.emp_company='"&pmg_company&"';"

		Set rs_emp = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rs_emp.EOF Or rs_emp.BOF Then
			emp_name = ""
		Else
			emp_no = rs_emp("emp_no")
			emp_name = f_toString(rs_emp("emp_name"), "")
			emp_company = f_toString(rs_emp("emp_company"), "")
			emp_bonbu = f_toString(rs_emp("emp_bonbu"), "")
			emp_saupbu = f_toString(rs_emp("emp_saupbu"), "")
			emp_team = f_toString(rs_emp("emp_team"), "")
			emp_org_code = f_toString(rs_emp("emp_org_code"), "")
			emp_org_name = f_toString(rs_emp("emp_org_name"), "")
			emp_reside_place = f_toString(rs_emp("emp_reside_place"), "")
			emp_reside_company = f_toString(rs_emp("emp_reside_company"), "")
			emp_in_date = f_toString(rs_emp("emp_in_date"), "")
			emp_grade = f_toString(rs_emp("emp_grade"), "")
			emp_position = f_toString(rs_emp("emp_position"), "")
			emp_type = f_toString(rs_emp("emp_type"), "")
			cost_center = f_toString(rs_emp("cost_center"), "")
			cost_group = f_toString(rs_emp("cost_group"), "")

			bank_name = f_toString(rs_emp("bank_name"), "")
			account_no = f_toString(rs_emp("account_no"), "")
			account_holder = f_toString(rs_emp("account_holder"), "")

			'//2017-06-09 비용구분(cost_group)이 잘못 등록된 경우(ex. 전사공통비, 전사공통비) 콤마 에러 이유가 됨. 앞부분만 등록처리
			If Trim(cost_center&"") <> "" And InStr(cost_center,",") > 0 Then
				cost_center = Left(cost_center, InStr(cost_center,","))
			End If
		End If

		w_cnt = w_cnt + 1

		' 지급항목
		pmg_base_pay = toString(xgr(4,i),0)	'기본급
		pmg_meals_pay = toString(xgr(5,i),0)	'식대
		pmg_postage_pay = toString(xgr(6,i),0)	'통신비(PL수당)
		pmg_re_pay = toString(xgr(7,i),0)	'소급급여
		pmg_overtime_pay = toString(xgr(8,i),0)	'연장근로수당

		'pmg_custom_pay	  = toString(xgr(20,i),"0")	'고객관리수당
		pmg_custom_pay = 0

		Select Case pmg_company
			Case "케이원"
				'지급항목
				pmg_car_pay = toString(xgr(9,i),0)	'주차지원금
				pmg_job_pay = toString(xgr(10,i),0)	'직무보조비(자격수당)
				pmg_job_support = toString(xgr(11,i) + xgr(15, i),0)	'업무장려비(업무장려비 + 시간외수당)
				pmg_jisa_pay = toString(xgr(12,i),0)	'본지사근무비
				pmg_disabled_pay = toString(xgr(13,i),0)	'장애인수당
				pmg_research_pay = toString(xgr(14,i),0)	'연구(연구수당)
				pmg_position_pay = toString(xgr(16,i),0)	'직책수당
				pmg_long_pay = toString(xgr(17,i),0)	'근속수당(PM수당)

				'공제항목
				de_nps_amt = toString(xgr(19,i),0)'국민연금
				de_nhis_amt = toString(xgr(20,i),0)'건강보험
				de_epi_amt = toString(xgr(21,i),0)'고용보험
				de_longcare_amt = toString(xgr(22,i),0)'장기요양보험료
				de_income_tax = toString(xgr(23,i),0)'소득세
				de_wetax = toString(xgr(24,i),0)'지방소득세
				de_year_incom_tax = toString(xgr(25,i),0)'연말정산소득세
				de_year_wetax = toString(xgr(26,i),0)'연말정산지방세
				de_other_amt1 = toString(xgr(30,i),0)'기타공제
				de_sawo_amt = toString(xgr(31,i),0)'사우회회비
				de_school_amt = toString(xgr(28,i) + xgr(35, i),0)'학자금대출(학자금상환+학자금대출)
				de_nhis_bla_amt = toString(xgr(33,i),0)'건강보험료정산
				de_long_bla_amt	= toString(xgr(34,i),0)'장기요양보험료정산
				de_hyubjo_amt = toString(xgr(32,i),0)'협조비

				de_year_incom_tax2 = toString(xgr(38,i),0)'연말재정산소득세
				de_year_wetax2 = toString(xgr(39,i),0)'연말재정산지방세
			Case "케이네트웍스"
				'지급항목
				pmg_car_pay = toString(xgr(9,i),0)	'주차지원금
				pmg_job_pay = toString(xgr(11,i),0)	'직무보조비(자격수당)
				pmg_job_support = toString(xgr(12,i) + xgr(14, i),0)	'업무장려비(업무장려비 + 시간외수당)
				pmg_jisa_pay = toString(xgr(13,i),0)	'본지사근무비
				pmg_disabled_pay = 0	'장애인수당
				pmg_research_pay = 0	'연구(연구수당)
				pmg_position_pay = toString(xgr(10,i),0)	'직책수당
				pmg_long_pay = toString(xgr(15,i),0)	'근속수당(PM수당)

				'공제항목
				de_nps_amt = toString(xgr(17,i),0)'국민연금
				de_nhis_amt = toString(xgr(18,i),0)'건강보험
				de_epi_amt = toString(xgr(19,i),0)'고용보험
				de_longcare_amt = toString(xgr(20,i),0)'장기요양보험료
				de_income_tax = toString(xgr(21,i),0)'소득세
				de_wetax = toString(xgr(22,i),0)'지방소득세
				de_year_incom_tax = toString(xgr(23,i),0)'연말정산소득세
				de_year_wetax = toString(xgr(24,i),0)'연말정산지방세

				de_other_amt1 = toString(xgr(27,i),0)'기타공제
				de_sawo_amt = toString(xgr(28,i),0)'사우회회비
				de_school_amt = toString(xgr(26,i) + xgr(31,i),0)'학자금대출(학자금상환+학자금대출)
				de_nhis_bla_amt = toString(xgr(29,i),0)'건강보험료정산
				de_long_bla_amt	= toString(xgr(30,i),0)'장기요양보험료정산
				de_hyubjo_amt = 0'협조비

				de_year_incom_tax2 = toString(xgr(35,i),0)'연말재정산소득세
				de_year_wetax2 = toString(xgr(36,i),0)'연말재정산지방세
			Case "케이시스템"
				'지급항목
				pmg_car_pay = 0
				pmg_job_pay = toString(xgr(11,i),0)	'직무보조비(자격수당)
				pmg_job_support = toString(xgr(9,i),0)	'업무장려비
				pmg_jisa_pay = 0	'본지사근무비
				pmg_disabled_pay = 0	'장애인수당
				pmg_research_pay = 0	'연구(연구수당)
				pmg_position_pay = 0	'직책수당
				pmg_long_pay = toString(xgr(10,i),0)	'근속수당(PM수당)

				'공제항목
				de_nps_amt = toString(xgr(13,i),0)'국민연금
				de_nhis_amt = toString(xgr(14,i),0)'건강보험
				de_epi_amt = toString(xgr(15,i),0)'고용보험
				de_longcare_amt = toString(xgr(16,i),0)'장기요양보험료
				de_income_tax = toString(xgr(17,i),0)'소득세
				de_wetax = toString(xgr(18,i),0)'지방소득세
				de_year_incom_tax = toString(xgr(19,i),0)'연말정산소득세
				de_year_wetax = toString(xgr(20,i),0)'연말정산지방세
				de_other_amt1 = toString(xgr(25,i),0)'기타공제
				de_sawo_amt = toString(xgr(26,i),0)'사우회회비
				de_school_amt = toString(xgr(22,i),0)'학자금상환
				de_nhis_bla_amt = toString(xgr(23,i),0)'건강보험료정산
				de_long_bla_amt	= toString(xgr(24,i),0)'장기요양보험료정산
				de_hyubjo_amt = 0'협조비

				de_year_incom_tax2 = toString(xgr(29,i),0)'연말재정산소득세
				de_year_wetax2 = toString(xgr(30,i),0)'연말재정산지방세
		End Select

		pmg_family_pay    = 0
		pmg_school_pay    = 0
		pmg_qual_pay      = 0
		pmg_other_pay1    = 0
		pmg_other_pay2    = 0
		pmg_other_pay3    = 0
		pmg_tax_yes       = 0
		pmg_tax_no        = 0
		pmg_tax_reduced   = 0

		pmg_give_total = pmg_base_pay + pmg_meals_pay + pmg_research_pay + pmg_postage_pay + pmg_re_pay
		pmg_give_total = pmg_give_total + pmg_overtime_pay + pmg_car_pay + pmg_position_pay + pmg_custom_pay + pmg_job_pay
		pmg_give_total = pmg_give_total + pmg_job_support + pmg_jisa_pay + pmg_long_pay + pmg_disabled_pay
		'pmg_give_total = xgr(25,i)

		meals_pay = pmg_meals_pay
		car_pay   = pmg_car_pay
		meals_tax_pay = 0
		car_tax_pay = 0

		If meals_pay > 100000 Then
			meals_tax_pay = meals_pay - 100000
			meals_pay =  100000
		End If

		If car_pay > 200000 Then
			car_tax_pay = car_pay - 200000
			car_pay = 200000
		End If

		pmg_tax_yes = pmg_give_total + meals_tax_pay + car_tax_pay
		pmg_tax_no = meals_pay + car_pay

		de_special_tax = 0
		de_saving_amt = 0
		de_johab_amt = 0

		de_deduct_total = de_nps_amt + de_nhis_amt + de_epi_amt + de_longcare_amt + de_income_tax
		de_deduct_total = de_deduct_total + de_wetax + de_year_incom_tax + de_year_wetax + de_year_incom_tax2 + de_year_wetax2
		de_deduct_total = de_deduct_total + de_other_amt1 + de_sawo_amt + de_school_amt + de_nhis_bla_amt + de_long_bla_amt
		de_deduct_total = de_deduct_total + de_hyubjo_amt
		'de_deduct_total = xgr(38,i)

		' 2019.03.15 윤성희,박정신 계산에 의한 공제액 합계가 아니라 엑셀 컬럼의 내용을 그대로 계산없이 설정한다.
		'de_deduct_total = xgr(43,i)

		'등록된 데이터가 있을 경우 삭제 후 입력 처리
		objBuilder.Append "SELECT pmg_emp_no FROM pay_month_give "
		objBuilder.Append "WHERE pmg_yymm = '"&pmg_yymm&"' AND pmg_id = '1' "
		objBuilder.Append "	AND pmg_emp_no = '"&emp_no&"' AND pmg_company = '"&pmg_company&"';"

		Set rs_give = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If Not rs_give.EOF Then
			objBuilder.Append "DELETE FROM pay_month_give "
			objBuilder.Append "WHERE pmg_yymm = '"&pmg_yymm&"' AND pmg_id = '1' "
			objBuilder.Append "	AND pmg_emp_no = '"&emp_no&"' AND pmg_company = '"&pmg_company&"';"

			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()
		End If
		rs_give.Close()

		objBuilder.Append "INSERT INTO pay_month_give(pmg_yymm, pmg_id, pmg_emp_no, pmg_company, pmg_date, "
		objBuilder.Append "pmg_in_date, pmg_emp_name, pmg_emp_type, pmg_org_code, pmg_org_name, "
		objBuilder.Append "pmg_bonbu, pmg_saupbu, pmg_team, pmg_reside_place, pmg_reside_company, "
		objBuilder.Append "pmg_grade, pmg_position, pmg_base_pay, pmg_meals_pay, pmg_postage_pay, "
		objBuilder.Append "pmg_re_pay, pmg_overtime_pay, pmg_car_pay, pmg_position_pay, pmg_custom_pay,  "
		objBuilder.Append "pmg_job_pay, pmg_job_support, pmg_jisa_pay, pmg_long_pay, pmg_disabled_pay, "
		objBuilder.Append "pmg_family_pay, pmg_school_pay, pmg_qual_pay, pmg_other_pay1, pmg_other_pay2, "
		objBuilder.Append "pmg_other_pay3, pmg_tax_yes, pmg_tax_no, pmg_tax_reduced, pmg_give_total, "
		objBuilder.Append "pmg_bank_name, pmg_account_no, pmg_account_holder, cost_group, cost_center, "
		objBuilder.Append "pmg_reg_date, pmg_reg_user, pmg_research_pay)"
		objBuilder.Append "VALUES('"&pmg_yymm&"', '1', '"&emp_no&"', '"&pmg_company&"', '"&pmg_date&"', "
		objBuilder.Append "'"&emp_in_date&"', '"&emp_name&"', '"&emp_type&"', '"&emp_org_code&"', '"&emp_org_name&"', "
		objBuilder.Append "'"&emp_bonbu&"', '"&emp_saupbu&"', '"&emp_team&"', '"&emp_reside_place&"', '"&emp_reside_company&"', "
		objBuilder.Append "'"&emp_grade&"','"&emp_position&"','"&pmg_base_pay&"','"&pmg_meals_pay&"','"&pmg_postage_pay&"', "
		objBuilder.Append "'"&pmg_re_pay&"', '"&pmg_overtime_pay&"', '"&pmg_car_pay&"', '"&pmg_position_pay&"', '"&pmg_custom_pay&"', "
		objBuilder.Append "'"&pmg_job_pay&"', '"&pmg_job_support&"', '"&pmg_jisa_pay&"', '"&pmg_long_pay&"', '"&pmg_disabled_pay&"', "
		objBuilder.Append "'"&pmg_family_pay&"', '"&pmg_school_pay&"', '"&pmg_qual_pay&"', '"&pmg_other_pay1&"', '"&pmg_other_pay2&"', "
		objBuilder.Append "'"&pmg_other_pay3&"', '"&pmg_tax_yes&"', '"&pmg_tax_no&"', '"&pmg_tax_reduced&"', '"&pmg_give_total&"', "
		objBuilder.Append "'"&bank_name&"', '"&account_no&"', '"&account_holder&"', '"&cost_group&"', '"&cost_center&"', "
		objBuilder.Append "NOW(),'"&user_name&"', '"&pmg_research_pay&"');"

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		objBuilder.Append "SELECT de_emp_no FROM pay_month_deduct "
		objBuilder.Append "WHERE de_yymm = '"&pmg_yymm&"' AND de_id = '1' "
		objBuilder.Append "	AND de_emp_no = '"&emp_no&"' AND de_company = '"&pmg_company&"';"

		Set rs_dct = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If Not rs_dct.EOF Then
			objBuilder.Append "DELETE FROM pay_month_deduct "
			objBuilder.Append "WHERE de_yymm = '"&pmg_yymm&"' AND de_id = '1' "
			objBuilder.Append "	AND de_emp_no = '"&emp_no&"' AND de_company = '"&pmg_company&"';"

			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()
		End If
		rs_dct.Close()

		objBuilder.Append "INSERT INTO pay_month_deduct(de_yymm, de_id, de_emp_no, de_company, de_date,"
		objBuilder.Append "de_emp_name, de_emp_type, de_org_code, de_org_name, de_bonbu,"
		objBuilder.Append "de_saupbu, de_team, de_reside_place, de_reside_company, de_grade,"
		objBuilder.Append "de_position, de_nps_amt, de_nhis_amt, de_epi_amt, de_longcare_amt,"
		objBuilder.Append "de_income_tax, de_wetax, de_year_incom_tax, de_year_wetax, de_year_incom_tax2,"
		objBuilder.Append "de_year_wetax2, de_other_amt1, de_saving_amt, de_sawo_amt, de_johab_amt,"
		objBuilder.Append "de_hyubjo_amt, de_school_amt, de_nhis_bla_amt, de_long_bla_amt, de_deduct_total,"
		objBuilder.Append "cost_group, cost_center, de_reg_date, de_reg_user)"
		objBuilder.Append "VALUES('"&pmg_yymm&"', '1', '"&emp_no&"', '"&pmg_company&"', '"&pmg_date&"',"
		objBuilder.Append "'"&emp_name&"', '"&emp_type&"', '"&emp_org_code&"', '"&emp_org_name&"', '"&emp_bonbu&"',"
		objBuilder.Append "'"&emp_saupbu&"', '"&emp_team&"', '"&emp_reside_place&"', '"&emp_reside_company&"', '"&emp_grade&"',"
		objBuilder.Append "'"&emp_position&"', '"&de_nps_amt&"', '"&de_nhis_amt&"', '"&de_epi_amt&"', '"&de_longcare_amt&"',"
		objBuilder.Append "'"&de_income_tax&"', '"&de_wetax&"', '"&de_year_incom_tax&"', '"&de_year_wetax&"','"&de_year_incom_tax2&"',"
		objBuilder.Append "'"&de_year_wetax2&"', '"&de_other_amt1&"', '"&de_saving_amt&"', '"&de_sawo_amt&"', '"&de_johab_amt&"',"
		objBuilder.Append "'"&de_hyubjo_amt&"', '"&de_school_amt&"', '"&de_nhis_bla_amt&"', '"&de_long_bla_amt&"', '"&de_deduct_total&"',"
		objBuilder.Append "'"&cost_group&"', '"&cost_center&"', NOW(), '"&user_name&"');"

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		'If Err.number <> 0 Then
		'	Response.Write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql4 & "]<br>"
		'End If
	Next

	Set rs_give = Nothing
	Set rs_dct = Nothing
End If


If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "등록 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = CStr(w_cnt)&" 건이 정상적으로 처리되었습니다."
End If

rs.Close : Set rs = Nothing
cn.close : set cn = Nothing

DBConn.Close() : Set DBConn = Nothing

'err_msg = cstr(rowcount+1) + " 건 처리되었습니다..."
Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	location.replace('/pay/insa_pay_month_pay_mg.asp');"
Response.Write "</script>"
Response.End
%>