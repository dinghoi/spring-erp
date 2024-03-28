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
Dim view_condi, pmg_yymm, pmg_emp_name
Dim curr_yyyy, curr_mm, title_line, savefilename
Dim sum_base_pay, sum_meals_pay, sum_postage_pay, sum_re_pay, sum_overtime_pay
Dim sum_car_pay, sum_position_pay, sum_custom_pay, sum_job_pay, sum_job_support
Dim sum_jisa_pay, sum_long_pay, sum_disabled_pay, sum_family_pay, sum_school_pay
Dim sum_qual_pay, sum_other_pay1, sum_other_pay2, sum_other_pay3, sum_tax_yes
Dim sum_tax_no, sum_tax_reduced, sum_give_tot, sum_nps_amt, sum_nhis_amt
Dim sum_epi_amt, sum_longcare_amt, sum_income_tax, sum_wetax, sum_year_incom_tax
Dim sum_year_wetax, sum_year_incom_tax2, sum_year_wetax2, sum_other_amt1, sum_sawo_amt
Dim sum_hyubjo_amt, sum_school_amt, sum_nhis_bla_amt, sum_long_bla_amt, sum_deduct_tot
Dim pay_count, sum_curr_pay

Dim sql

view_condi = Request.QueryString("view_condi")
pmg_yymm = Request.QueryString("pmg_yymm")
pmg_emp_name = Request.QueryString("pmg_emp_name")

curr_yyyy = Mid(CStr(pmg_yymm), 1, 4)
curr_mm = Mid(CStr(pmg_yymm), 5, 2)
title_line = CStr(curr_yyyy) & "�� " & CStr(curr_mm) & "�� �޿��̿� ������(���κ�)"

savefilename = title_line & ".xls"

Call ViewExcelType(savefilename)

'===================================================
'### DB Query & Call Procedure
'===================================================
Dim objCmd, objRs

Set objCmd = Server.CreateObject("ADODB.Command")
With objCmd
    .ActiveConnection = DBConn
    .CommandText = "USP_PAY_INSA_PAY_EXCEL_PAY_PAY_REPORT_SEL"
    .CommandType = adCmdStoredProc

    .Parameters.Append .CreateParameter("p_pmg_yymm", adVarChar, adParamInput, 6, pmg_yymm)
	.Parameters.Append .CreateParameter("p_emp_company", adVarChar, adParamInput, 6, view_condi)
	.Parameters.Append .CreateParameter("p_pmg_emp_name", adVarChar, adParamInput, 20, pmg_emp_name)

	Set objRs = .Execute()
End With

Set objCmd = Nothing

'===================================================
sum_base_pay = 0 : sum_meals_pay = 0 : sum_postage_pay = 0 : sum_re_pay = 0 : sum_overtime_pay = 0
sum_car_pay = 0 : sum_position_pay = 0 : sum_custom_pay = 0 : sum_job_pay = 0 : sum_job_support = 0
sum_jisa_pay = 0 : sum_long_pay = 0 : sum_disabled_pay = 0 : sum_family_pay = 0 : sum_school_pay = 0
sum_qual_pay = 0 : sum_other_pay1 = 0 : sum_other_pay2 = 0 : sum_other_pay3 = 0 : sum_tax_yes = 0
sum_tax_no = 0 : sum_tax_reduced = 0 : sum_give_tot = 0 : sum_nps_amt = 0 : sum_nhis_amt = 0
sum_epi_amt = 0 : sum_longcare_amt = 0 : sum_income_tax = 0 : sum_wetax = 0 : sum_year_incom_tax = 0
sum_year_wetax = 0 : sum_year_incom_tax2 = 0 : sum_year_wetax2 = 0 : sum_other_amt1 = 0 : sum_sawo_amt = 0
sum_hyubjo_amt = 0 : sum_school_amt = 0 : sum_nhis_bla_amt = 0 : sum_long_bla_amt = 0 : sum_deduct_tot = 0

pay_count = 0
sum_curr_pay = 0

'SQL = "SELECT pmgt.pmg_emp_no, pmgt.pmg_company, pmgt.pmg_give_total, pmgt.pmg_base_pay, pmgt.pmg_meals_pay, "
'SQL = SQL & "	pmgt.pmg_postage_pay, pmgt.pmg_re_pay, pmgt.pmg_overtime_pay, pmgt.pmg_car_pay, pmgt.pmg_position_pay, "
'SQL = SQL & "	pmgt.pmg_custom_pay, pmgt.pmg_job_pay, pmgt.pmg_job_support, pmgt.pmg_jisa_pay, pmgt.pmg_long_pay, "
'SQL = SQL & "	pmgt.pmg_disabled_pay, pmgt.pmg_give_total, pmgt.pmg_emp_name, pmgt.pmg_in_date, pmgt.pmg_grade, "
'SQL = SQL & "	pmgt.pmg_bonbu, pmgt.pmg_saupbu, pmgt.pmg_team, pmgt.pmg_org_name, pmgt.pmg_reside_place, "
'SQL = SQL & "	pmgt.pmg_reside_company, pmgt.cost_group, pmgt.cost_center, "
'SQL = SQL & "	eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, "
'SQL = SQL & "	eomt.org_reside_place, eomt.org_reside_company, emmt.cost_group AS costGroup, emmt.cost_center AS costCenter, "
'SQL = SQL &  "	pmdt.de_nps_amt, pmdt.de_nhis_amt, pmdt.de_epi_amt, pmdt.de_longcare_amt, pmdt.de_income_tax, pmdt.de_wetax, "
'SQL = SQL & "	pmdt.de_year_incom_tax, pmdt.de_year_wetax, pmdt.de_year_incom_tax2, pmdt.de_year_wetax2, pmdt.de_other_amt1, "
'SQL = SQL & "	pmdt.de_sawo_amt, pmdt.de_hyubjo_amt, pmdt.de_school_amt, pmdt.de_nhis_bla_amt, pmdt.de_long_bla_amt, pmdt.de_deduct_total "
'SQL = SQL & "FROM pay_month_give AS pmgt "
'SQL = SQL & "INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no "
'SQL = SQL & "	AND emmt.emp_month = '"&pmg_yymm&"' "
'SQL = SQL & "INNER JOIN emp_org_mst AS eomt ON emmt.emp_org_code = eomt.org_code "
'SQL = SQL & "INNER JOIN pay_month_deduct AS pmdt ON pmgt.pmg_emp_no = pmdt.de_emp_no "
'SQL = SQL & "WHERE pmgt.pmg_yymm = '"&pmg_yymm&"' AND pmgt.pmg_id = '1' "
'SQL = SQL & "	AND pmdt.de_yymm = '"&pmg_yymm&"' AND pmdt.de_id = '1' "
'SQL = SQL & "	AND eomt.org_company = '"&view_condi&"' "
'SQL = SQL & " AND pmgt.pmg_emp_name LIKE '%" & pmg_emp_name & "%' "
'SQL = SQL & " ORDER BY pmgt.pmg_company, pmgt.pmg_org_code, pmgt.pmg_emp_no ASC"

'Set Rs = Server.CreateObject("ADODB.Recordset")
'Rs.Open SQL, Dbconn, 1

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
	<tr>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">�ͼӳ��</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">���</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">��  ��</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">�Ի���</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">����</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">ȸ��</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">����</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">�����</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">��</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">�μ�</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">����ó</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">����óȸ��</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">��뼾Ÿ�׷�</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">��뱸��</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�⺻��</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�Ĵ�</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">��ź�</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�ұޱ޿�</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">����ٷμ���</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">����������</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">��å����</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">����������</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">����������</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">���������</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">������ٹ���</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�ټӼ���</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">����μ���</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�����հ�</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">���ο���</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�ǰ�����</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">��뺸��</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">����纸���</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�ҵ漼</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">����ҵ漼</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">��������ҵ漼</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�����������漼</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">����������ҵ漼</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�������������漼</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">��Ÿ����</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">���ȸ ȸ��</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">���ڱݻ�ȯ</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�ǰ����������</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">����纸�������</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">������</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�����հ�</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�������޾�</div></td>
	</tr>
	<%
	Dim pmg_give_tot, de_deduct_tot, pmg_curr_pay
	Dim de_nps_amt, de_nhis_amt, de_epi_amt, de_longcare_amt, de_income_tax
	Dim de_wetax, de_year_incom_tax, de_year_wetax, de_year_incom_tax2, de_year_wetax2
	Dim de_other_amt1, de_sawo_amt, de_hyubjo_amt, de_school_amt, de_nhis_bla_amt, de_long_bla_amt

	Do Until objRs.EOF
		'emp_no = objRs("pmg_emp_no")
		'pmg_company = objRs("pmg_company")
		pmg_give_tot = objRs("pmg_give_total")
		pay_count = pay_count + 1

		sum_base_pay = sum_base_pay + Int(objRs("pmg_base_pay"))
		sum_meals_pay = sum_meals_pay + Int(objRs("pmg_meals_pay"))
		sum_postage_pay = sum_postage_pay + Int(objRs("pmg_postage_pay"))
		sum_re_pay = sum_re_pay + Int(objRs("pmg_re_pay"))
		sum_overtime_pay = sum_overtime_pay + Int(objRs("pmg_overtime_pay"))
		sum_car_pay = sum_car_pay + Int(objRs("pmg_car_pay"))
		sum_position_pay = sum_position_pay + Int(objRs("pmg_position_pay"))
		sum_custom_pay = sum_custom_pay + Int(objRs("pmg_custom_pay"))
		sum_job_pay = sum_job_pay + Int(objRs("pmg_job_pay"))
		sum_job_support = sum_job_support + Int(objRs("pmg_job_support"))
		sum_jisa_pay = sum_jisa_pay + Int(objRs("pmg_jisa_pay"))
		sum_long_pay = sum_long_pay + Int(objRs("pmg_long_pay"))
		sum_disabled_pay = sum_disabled_pay + Int(objRs("pmg_disabled_pay"))
		sum_give_tot = sum_give_tot + Int(objRs("pmg_give_total"))



		de_nps_amt = Int(objRs("de_nps_amt"))
		de_nhis_amt = Int(objRs("de_nhis_amt"))
		de_epi_amt = Int(objRs("de_epi_amt"))
		de_longcare_amt = Int(objRs("de_longcare_amt"))
		de_income_tax = Int(objRs("de_income_tax"))
		de_wetax = Int(objRs("de_wetax"))
		de_year_incom_tax = Int(objRs("de_year_incom_tax"))
		de_year_wetax = Int(objRs("de_year_wetax"))
		de_year_incom_tax2 = Int(objRs("de_year_incom_tax2"))
		de_year_wetax2 = Int(objRs("de_year_wetax2"))
		de_other_amt1 = Int(objRs("de_other_amt1"))
		de_sawo_amt = Int(objRs("de_sawo_amt"))
		de_hyubjo_amt = Int(objRs("de_hyubjo_amt"))
		de_school_amt = Int(objRs("de_school_amt"))
		de_nhis_bla_amt = Int(objRs("de_nhis_bla_amt"))
		de_long_bla_amt = Int(objRs("de_long_bla_amt"))
		de_deduct_tot = Int(objRs("de_deduct_total"))

		pmg_curr_pay = pmg_give_tot - de_deduct_tot

		sum_nps_amt = sum_nps_amt + de_nps_amt
		sum_nhis_amt = sum_nhis_amt + de_nhis_amt
		sum_epi_amt = sum_epi_amt + de_epi_amt
		sum_longcare_amt = sum_longcare_amt + de_longcare_amt
		sum_income_tax = sum_income_tax + de_income_tax
		sum_wetax = sum_wetax + de_wetax
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
	%>
  <tr valign="middle" class="style11">
    <td width="110"><div align="center" class="style1"><%=pmg_yymm%></div></td>
    <td width="110"><div align="center" class="style1"><%=objRs("pmg_emp_no")%></div></td>
    <td width="110"><div align="center" class="style1"><%=objRs("pmg_emp_name")%></div></td>
    <td width="110"><div align="center" class="style1"><%=objRs("pmg_in_date")%></div></td>
    <td width="110"><div align="center" class="style1"><%=objRs("pmg_grade")%></div></td>
    <td width="110"><div align="center" class="style1"><%=objRs("org_company")%></div></td>
    <td width="110"><div align="center" class="style1"><%=objRs("org_bonbu")%></div></td>
    <td width="110"><div align="center" class="style1"><%=objRs("org_saupbu")%></div></td>
    <td width="110"><div align="center" class="style1"><%=objRs("org_team")%></div></td>
    <td width="110"><div align="center" class="style1"><%=objRs("org_name")%></div></td>
    <td width="110"><div align="center" class="style1"><%=objRs("org_reside_place")%></div></td>
    <td width="110"><div align="center" class="style1"><%=objRs("org_reside_company")%></div></td>
    <td width="110"><div align="center" class="style1"><%=objRs("costGroup")%></div></td>
    <td width="110"><div align="center" class="style1"><%=objRs("costCenter")%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(objRs("pmg_base_pay"), 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(objRs("pmg_meals_pay"), 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(objRs("pmg_postage_pay"), 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(objRs("pmg_re_pay"), 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(objRs("pmg_overtime_pay"), 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(objRs("pmg_car_pay"), 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(objRs("pmg_position_pay"), 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(objRs("pmg_custom_pay"), 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(objRs("pmg_job_pay"), 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(objRs("pmg_job_support"), 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(objRs("pmg_jisa_pay"), 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(objRs("pmg_long_pay"), 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(objRs("pmg_disabled_pay"), 0)%></div></td>
    <td width="100"><div align="right" class=","><%=Formatnumber(objRs("pmg_give_total"), 0)%></div></td>
    <%


    %>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(de_nps_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(de_nhis_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(de_epi_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(de_longcare_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(de_income_tax, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(de_wetax, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(de_year_incom_tax, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(de_year_wetax, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(de_year_incom_tax2, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(de_year_wetax2, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(de_other_amt1, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(de_sawo_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(de_school_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(de_nhis_bla_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(de_long_bla_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(de_hyubjo_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(de_deduct_tot, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(pmg_curr_pay, 0)%></div></td>
  </tr>
	<%
		objRs.MoveNext()
	Loop
	objRs.Close() : Set objRs = Nothing
	DBConn.Close() : Set DBConn = Nothing

	sum_curr_pay = sum_give_tot - sum_deduct_tot
	%>
  <tr valign="middle" class="style11">
    <td colspan="13" width="110"><div align="center" class="style1">�Ѱ�</div></td>
    <td width="110"><div align="center" class="style1"><%=Formatnumber(pay_count, 0)%>&nbsp;��</div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_base_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_meals_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_postage_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_re_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_overtime_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_car_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_position_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_custom_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_job_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_job_support, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_jisa_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_long_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_disabled_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_give_tot, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_nps_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_nhis_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_epi_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_longcare_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_income_tax, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_wetax, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_year_incom_tax, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_year_wetax, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_year_incom_tax2, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_year_wetax2, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_other_amt1, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_sawo_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_school_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_nhis_bla_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_long_bla_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_hyubjo_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_deduct_tot, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=Formatnumber(sum_curr_pay, 0)%></div></td>
  </tr>
</table>
</body>
</html>