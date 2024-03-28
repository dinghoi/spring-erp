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
Dim view_condi, pmg_yymm, to_date, give_date, curr_yyyy, curr_mm, title_line
Dim savefilename, sum_base_pay, sum_meals_pay, sum_postage_pay, sum_re_pay, sum_overtime_pay
Dim sum_car_pay, sum_position_pay, sum_custom_pay, sum_job_pay, sum_job_support
Dim sum_jisa_pay, sum_long_pay, sum_disabled_pay, sum_family_pay, sum_school_pay
Dim sum_qual_pay, sum_other_pay1, sum_other_pay2, sum_other_pay3, sum_tax_yes
Dim sum_tax_no, sum_tax_reduced, sum_give_tot, sum_nps_amt, sum_nhis_amt
Dim sum_epi_amt, sum_longcare_amt, sum_income_tax, sum_wetax, sum_year_incom_tax
Dim sum_year_wetax, sum_year_incom_tax2, sum_year_wetax2, sum_other_amt1, sum_sawo_amt
Dim sum_hyubjo_amt, sum_school_amt, sum_nhis_bla_amt, sum_long_bla_amt, sum_deduct_tot
Dim pay_count, sum_curr_pay, rsPay

Dim emp_first_date, emp_in_date, emp_end_date, emp_bonbu, emp_saupbu, emp_team
Dim pmg_emp_no, pmg_give_tot, de_nps_amt, de_nhis_amt, de_epi_amt, de_longcare_amt, de_income_tax
Dim de_wetax, de_year_incom_tax, de_year_wetax, de_year_incom_tax2, de_year_wetax2, de_other_amt1
Dim de_sawo_amt, de_hyubjo_amt, de_school_amt, de_nhis_bla_amt, de_long_bla_amt
Dim de_deduct_tot, pmg_curr_pay

view_condi = Request.QueryString("view_condi")
pmg_yymm = Request.QueryString("pmg_yymm")
to_date = Request.QueryString("to_date")

'curr_date = datevalue(mid(cstr(now()),1,10))

give_date = to_date '지급일

curr_yyyy = Mid(CStr(pmg_yymm),1,4)
curr_mm = Mid(CStr(pmg_yymm),5,2)
title_line = CStr(curr_yyyy)&"년 "&CStr(curr_mm)&"월 "&" 급여대장("&view_condi&")"

savefilename = title_line&".xls"
'savefilename = "입사자 현황 -- "+ to_date +""+ view_condi +"" + cstr(curr_date) + ".xls"

'엑셀 타입 형식
Call ViewExcelType(savefilename)

sum_base_pay = 0
sum_meals_pay = 0
sum_postage_pay = 0
sum_re_pay = 0
sum_overtime_pay = 0
sum_car_pay = 0
sum_position_pay = 0
sum_custom_pay = 0
sum_job_pay = 0
sum_job_support = 0
sum_jisa_pay = 0
sum_long_pay = 0
sum_disabled_pay = 0
sum_family_pay = 0
sum_school_pay = 0
sum_qual_pay = 0
sum_other_pay1 = 0
sum_other_pay2 = 0
sum_other_pay3 = 0
sum_tax_yes = 0
sum_tax_no = 0
sum_tax_reduced = 0
sum_give_tot = 0
sum_nps_amt = 0
sum_nhis_amt = 0
sum_epi_amt = 0
sum_longcare_amt = 0
sum_income_tax = 0
sum_wetax = 0
sum_year_incom_tax = 0
sum_year_wetax = 0
sum_year_incom_tax2 = 0
sum_year_wetax2 = 0
sum_other_amt1 = 0
sum_sawo_amt = 0
sum_hyubjo_amt = 0
sum_school_amt = 0
sum_nhis_bla_amt = 0
sum_long_bla_amt = 0
sum_deduct_tot = 0

pay_count = 0
sum_curr_pay = 0
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
<!--
.style1 {font-size: 12px}
.style2 {
	font-size: 14px;
	font-weight: bold;
}
-->
</style>
</head>
<body>
<table  border="0" cellpadding="0" cellspacing="0">
	<tr bgcolor="#EFEFEF" class="style11">
		<td colspan="16" bgcolor="#FFFFFF"><div align="left" class="style2"><%=title_line%></div></td>
	</tr>
	<tr bgcolor="#EFEFEF" class="style11">
		<td colspan="2" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">인적사항</div></td>
		<td colspan="7" style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><div align="center" class="style1">기본급여 및 제수당</div></td>
		<td colspan="6" style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">공제 및 차인지급액</div></td>
	</tr>
	<tr>
		<td style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><div align="center" class="style1">사번</div></td>
		<td style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><div align="center" class="style1">성  명</div></td>
		<td><div align="center" class="style1">기본급</div></td>
		<td><div align="center" class="style1">식대</div></td>
		<td><div align="center" class="style1">차량유지비</div></td>
		<td><div align="center" class="style1">통신비</div></td>
		<td><div align="center" class="style1">소급급여</div></td>
		<td><div align="center" class="style1">연장근로수당</div></td>
		<td><div align="center" class="style1">주차지원금</div></td>
		<td><div align="center" class="style1">국민연금</div></td>
		<td><div align="center" class="style1">건강보험</div></td>
		<td><div align="center" class="style1">고용보험</div></td>
		<td><div align="center" class="style1">장기요양보험료</div></td>
		<td><div align="center" class="style1">소득세</div></td>
		<td><div align="center" class="style1">지방소득세</div></td>
	</tr>
	<tr>
		<td style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><div align="center" class="style1">입사일</div></td>
		<td style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><div align="center" class="style1">직급</div></td>
		<td><div align="center" class="style1">직책수당</div></td>
		<td><div align="center" class="style1">고객관리수당</div></td>
		<td><div align="center" class="style1">직무보조비</div></td>
		<td><div align="center" class="style1">업무장려비</div></td>
		<td><div align="center" class="style1">본지사근무비</div></td>
		<td><div align="center" class="style1">근속수당</div></td>
		<td><div align="center" class="style1">장애인수당</div></td>
		<td><div align="center" class="style1">기타공제</div></td>
		<td><div align="center" class="style1">사우회 회비</div></td>
		<td><div align="center" class="style1">학자금상환</div></td>
		<td><div align="center" class="style1">건강보험료정산</div></td>
		<td><div align="center" class="style1">장기요양보험료정산</div></td>
		<td><div align="center" class="style1">공제합계</div></td>
	</tr>
	<tr>
		<td style=" border-bottom:2px solid #515254; background:#f8f8f8;"><div align="center" class="style1">퇴사일</div></td>
		<td style=" border-bottom:2px solid #515254; background:#f8f8f8;"><div align="center" class="style1">부서</div></td>
		<td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">&nbsp;</div></td>
		<td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">&nbsp;</div></td>
		<td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">&nbsp;</div></td>
		<td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">&nbsp;</div></td>
		<td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">&nbsp;</div></td>
		<td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">&nbsp;</div></td>
		<td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">지급합계</div></td>
		<td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">협조비</div></td>
		<td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">연말정산소득세</div></td>
		<td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">연말정산지방세</div></td>
		<td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">연말재정산소득세</div></td>
		<td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">연말재정산지방세</div></td>
		<td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">차인지급액</div></td>
	</tr>
	<%
	'Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_bonbu,pmg_saupbu,pmg_team,pmg_org_name,pmg_emp_no ASC"
	objBuilder.Append "SELECT pmgt.pmg_emp_no, pmgt.pmg_give_total, pmgt.pmg_emp_name, pmgt.pmg_grade, "
	objBuilder.Append " pmgt.pmg_org_name, pmgt.pmg_bank_name, pmgt.pmg_account_no, pmgt.pmg_account_holder, "
	objBuilder.Append "	pmgt.pmg_base_pay, pmgt.pmg_meals_pay, pmgt.pmg_postage_pay, pmgt. pmg_re_pay, "
	objBuilder.Append "	pmgt.pmg_overtime_pay, pmgt.pmg_car_pay, pmgt.pmg_position_pay, pmgt.pmg_custom_pay, "
	objBuilder.Append "	pmgt.pmg_job_pay, pmgt.pmg_job_support, pmgt.pmg_jisa_pay, pmgt.pmg_long_pay, "
	objBuilder.Append "	pmgt.pmg_disabled_pay, "

	objBuilder.Append "	emtt.emp_in_date, emtt.emp_jikmu, emtt.emp_first_date, emtt.emp_end_date, emtt.emp_company, "
	objBuilder.Append "	emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team, "

	objBuilder.Append "	pmdt.de_nps_amt, pmdt.de_nhis_amt, pmdt.de_epi_amt, pmdt.de_longcare_amt, pmdt.de_income_tax, "
	objBuilder.Append "	pmdt.de_wetax, pmdt.de_year_incom_tax, pmdt.de_year_wetax, pmdt.de_year_incom_tax2, "
	objBuilder.Append "	pmdt.de_year_wetax2, pmdt.de_other_amt1, pmdt.de_sawo_amt, pmdt.de_hyubjo_amt, "
	objBuilder.Append "	pmdt.de_year_wetax2, pmdt.de_other_amt1, pmdt.de_sawo_amt, pmdt.de_hyubjo_amt, "
	objBuilder.Append "	pmdt.de_school_amt, pmdt.de_nhis_bla_amt, pmdt.de_long_bla_amt, pmdt.de_deduct_total, "

	objBuilder.Append "	(SELECT SUM(pmg_give_total) FROM pay_month_give "
	objBuilder.Append "	WHERE pmg_yymm = pmgt.pmg_yymm AND pmg_id = '1' AND pmg_company = pmgt.pmg_company) AS 'pmg_give_tot', "
	objBuilder.Append "	(SELECT SUM(de_deduct_total) FROM pay_month_deduct "
	objBuilder.Append "	WHERE de_yymm = pmgt.pmg_yymm AND de_id = '1' AND de_company = pmgt.pmg_company) AS 'de_deduct_tot' "
	objBuilder.Append "FROM pay_month_give AS pmgt "
	objBuilder.Append "INNER JOIN emp_master AS emtt ON pmgt.pmg_emp_no = emtt.emp_no "
	objBuilder.Append "	AND (ISNULL(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' Or emtt.emp_end_date = '') "
	objBuilder.Append "LEFT OUTER JOIN pay_month_deduct AS pmdt ON pmgt.pmg_emp_no = pmdt.de_emp_no "
	objBuilder.Append "	AND pmgt.pmg_company = pmdt.de_company AND pmdt.de_id = '1' AND de_yymm = '"&pmg_yymm&"' "
	objBuilder.Append "WHERE pmg_yymm = '"&pmg_yymm&"' AND pmg_id = '1' AND pmg_company = '"&view_condi&"' "
	objBuilder.Append "ORDER BY pmgt.pmg_company, pmgt.pmg_bank_name, pmgt.pmg_emp_no ASC;"

	Set rsPay = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	Do Until rsPay.EOF
		pay_count = pay_count + 1

		pmg_emp_no = rsPay("pmg_emp_no")
		pmg_give_tot = rsPay("pmg_give_total")

		emp_in_date = rsPay("emp_in_date")
		emp_end_date = rsPay("emp_end_date")

		sum_base_pay = sum_base_pay + int(rsPay("pmg_base_pay"))
		sum_meals_pay = sum_meals_pay + int(rsPay("pmg_meals_pay"))
		sum_postage_pay = sum_postage_pay + int(rsPay("pmg_postage_pay"))
		sum_re_pay = sum_re_pay + int(rsPay("pmg_re_pay"))
		sum_overtime_pay = sum_overtime_pay + int(rsPay("pmg_overtime_pay"))
		sum_car_pay = sum_car_pay + int(rsPay("pmg_car_pay"))
		sum_position_pay = sum_position_pay + int(rsPay("pmg_position_pay"))
		sum_custom_pay = sum_custom_pay + int(rsPay("pmg_custom_pay"))
		sum_job_pay = sum_job_pay + int(rsPay("pmg_job_pay"))
		sum_job_support = sum_job_support + int(rsPay("pmg_job_support"))
		sum_jisa_pay = sum_jisa_pay + int(rsPay("pmg_jisa_pay"))
		sum_long_pay = sum_long_pay + int(rsPay("pmg_long_pay"))
		sum_disabled_pay = sum_disabled_pay + int(rsPay("pmg_disabled_pay"))
		sum_give_tot = sum_give_tot + int(rsPay("pmg_give_total"))

		de_nps_amt = int(rsPay("de_nps_amt"))
		de_nhis_amt = int(rsPay("de_nhis_amt"))
		de_epi_amt = int(rsPay("de_epi_amt"))
		de_longcare_amt = int(rsPay("de_longcare_amt"))
		de_income_tax = int(rsPay("de_income_tax"))
		de_wetax = int(rsPay("de_wetax"))
		de_year_incom_tax = int(rsPay("de_year_incom_tax"))
		de_year_wetax = int(rsPay("de_year_wetax"))
		de_year_incom_tax2 = int(rsPay("de_year_incom_tax2"))
		de_year_wetax2 = int(rsPay("de_year_wetax2"))
		de_other_amt1 = int(rsPay("de_other_amt1"))
		de_sawo_amt = int(rsPay("de_sawo_amt"))
		de_hyubjo_amt = int(rsPay("de_hyubjo_amt"))
		de_school_amt = int(rsPay("de_school_amt"))
		de_nhis_bla_amt = int(rsPay("de_nhis_bla_amt"))
		de_long_bla_amt = int(rsPay("de_long_bla_amt"))
		de_deduct_tot = int(rsPay("de_deduct_total"))

		sum_nps_amt = sum_nps_amt + de_nps_amt
		sum_nhis_amt = sum_nhis_amt + de_nhis_amt
		sum_epi_amt = sum_epi_amt + de_epi_amt
		sum_longcare_amt = sum_longcare_amt + de_longcare_amt
		sum_income_tax = sum_income_tax + de_income_tax
		sum_wetax = sum_wetax + de_wetax
		sum_year_incom_tax = sum_year_incom_tax + de_year_incom_tax
		sum_year_wetax = sum_year_wetax + de_year_wetax
		sum_year_incom_tax2 = sum_year_incom_tax2 + de_year_incom_tax2
		sum_year_wetax2 = sum_year_wetax2 + de_year_wetax2
		sum_other_amt1 = sum_other_amt1 + de_other_amt1
		sum_sawo_amt = sum_sawo_amt + de_sawo_amt
		sum_hyubjo_amt = sum_hyubjo_amt + de_hyubjo_amt
		sum_school_amt = sum_school_amt + de_school_amt
		sum_nhis_bla_amt = sum_nhis_bla_amt + de_nhis_bla_amt
		sum_long_bla_amt = sum_long_bla_amt + de_long_bla_amt
		sum_deduct_tot = sum_deduct_tot + de_deduct_tot

		pmg_curr_pay = pmg_give_tot - de_deduct_tot
	%>
	<tr valign="middle" class="style11" <%If pay_count Mod 2 = 0 Then %>style="background-color:#EAEAEA;"<%End If%>>
		<td width="110"><div align="center" class="style1"><%=rsPay("pmg_emp_no")%></div></td>
		<td width="110"><div align="center" class="style1"><%=rsPay("pmg_emp_name")%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(rsPay("pmg_base_pay"),0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(rsPay("pmg_meals_pay"),0)%></div></td>
		<td width="100"><div align="right" class="style1">&nbsp;</div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(rsPay("pmg_postage_pay"),0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(rsPay("pmg_re_pay"),0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(rsPay("pmg_overtime_pay"),0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(rsPay("pmg_car_pay"),0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(de_nps_amt,0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(de_nhis_amt,0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(de_epi_amt,0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(de_longcare_amt,0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(de_income_tax,0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(de_wetax,0)%></div></td>
	</tr>
	<tr <%If pay_count Mod 2 = 0 Then %>style="background-color:#EAEAEA;"<%End If%>>
		<td width="110"><div align="center" class="style1"><%=emp_in_date%></div></td>
		<td width="110"><div align="center" class="style1"><%=rsPay("pmg_grade")%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(rsPay("pmg_position_pay"),0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(rsPay("pmg_custom_pay"),0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(rsPay("pmg_job_pay"),0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(rsPay("pmg_job_support"),0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(rsPay("pmg_jisa_pay"),0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(rsPay("pmg_long_pay"),0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(rsPay("pmg_disabled_pay"),0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(de_other_amt1,0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(de_sawo_amt,0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(de_school_amt,0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(de_nhis_bla_amt,0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(de_long_bla_amt,0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=formatnumber(de_deduct_tot,0)%></div></td>
	</tr>
	<tr <%If pay_count Mod 2 = 0 Then %>style="background-color:#EAEAEA;"<%End If%>>
		<td width="110" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1"><%=emp_end_date%></div></td>
		<td width="110" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1"><%=rsPay("pmg_org_name")%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">&nbsp;</div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">&nbsp;</div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">&nbsp;</div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">&nbsp;</div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">&nbsp;</div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">&nbsp;</div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(rsPay("pmg_give_total"),0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(de_hyubjo_amt,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(de_year_incom_tax,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(de_year_wetax,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(de_year_incom_tax2,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(de_year_wetax2,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(pmg_curr_pay,0)%></div></td>
	</tr>
	<%
		rsPay.MoveNext()
	Loop
	rsPay.Close() : Set rsPay = Nothing
	DBConn.Close() : Set DBConn = Nothing

	sum_curr_pay = sum_give_tot - sum_deduct_tot
	%>
	<tr>
		<th rowspan="3" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="center" class="style1">총계</div></th>
		<th rowspan="3" width="110" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(pay_count,0)%>&nbsp;명</div></th>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_base_pay,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_meals_pay,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1">&nbsp;</div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_postage_pay,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_re_pay,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_overtime_pay,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_car_pay,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_nps_amt,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_nhis_amt,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_epi_amt,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_longcare_amt,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_income_tax,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_wetax,0)%></div></td>
	</tr>
	<tr>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_position_pay,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_custom_pay,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_job_pay,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_job_support,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_jisa_pay,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_long_pay,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_disabled_pay,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_other_amt1,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_sawo_amt,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_school_amt,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_nhis_bla_amt,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_long_bla_amt,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_deduct_tot,0)%></div></td>
	</tr>
	<tr>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1">&nbsp;</div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1">&nbsp;</div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1">&nbsp;</div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1">&nbsp;</div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1">&nbsp;</div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1">&nbsp;</div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_give_tot,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_hyubjo_amt,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_year_incom_tax,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_year_wetax,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_year_incom_tax2,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_year_wetax2,0)%></div></td>
		<td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_curr_pay,0)%></div></td>
	</tr>
</table>
</body>
</html>