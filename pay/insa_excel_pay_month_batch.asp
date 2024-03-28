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
Dim view_condi, pmg_yymm, pmg_yymm_to, to_date, v_company
Dim give_date, curr_yyyy, curr_mm, title_line, savefilename
Dim st_in_date, rever_year
Dim rsInsEmp, rsInsHap, rsPay, epi_emp, epi_com, long_hap

Dim emp_name, emp_in_date, pmg_grade, pmg_company
Dim pmg_bonbu, pmg_saupbu, pmg_team, pmg_org_name
Dim pmg_base_pay, pmg_meals_pay, pmg_postage_pay, pmg_re_pay
Dim pmg_overtime_pay, pmg_car_pay, pmg_position_pay, pmg_custom_pay, pmg_job_pay
Dim pmg_job_support, pmg_jisa_pay, pmg_long_pay, pmg_disabled_pay, pmg_give_total
Dim de_nps_amt, de_nhis_amt, de_epi_amt, de_longcare_amt, de_income_tax, de_wetax
Dim de_year_incom_tax, de_year_wetax, de_year_incom_tax2, de_year_wetax2, de_other_amt1
Dim de_sawo_amt, de_hyubjo_amt, de_school_amt, de_nhis_bla_amt, de_long_bla_amt
Dim de_deduct_tot, incom_family_cnt, pmg_curr_pay, incom_month_amount
Dim incom_nps, incom_nhis, incom_go_yn, incom_long_yn, incom_wife_yn
Dim incom_age20, incom_age60, incom_old, pmg_tax_yes, pmg_tax_no, inc_st_amt, inc_incom
Dim rs_sod, long_amt, pmg_give_tot, epi_amt, we_tax
Dim pmg_emp_no, de_emp_no, incom_emp_no, incom_base_pay
Dim incom_overtime_pay

Dim arrPay, i

view_condi = Request.QueryString("view_condi")
pmg_yymm = Request.QueryString("pmg_yymm")
pmg_yymm_to = Request.QueryString("pmg_yymm_to")
to_date = Request.QueryString("to_date")

'curr_date = datevalue(mid(cstr(now()),1,10))

'if view_condi = "����������ġ" then
'	v_company = "�ڸ��Ƶ𿣾�"
'else
'	v_company = view_condi
'end if

give_date = to_date '������

curr_yyyy = Mid(CStr(pmg_yymm), 1, 4)
curr_mm = Mid(CStr(pmg_yymm), 5, 2)
title_line = CStr(curr_yyyy)& "�� "&CStr(curr_mm)&"�� "&" �޿��̿� ������(���κ�)-"&view_condi

savefilename = title_line &".xls"
'savefilename = "�Ի��� ��Ȳ -- "+ to_date +""+ view_condi +"" + cstr(curr_date) + ".xls"

Call ViewExcelType(savefilename)

'��� �Ի�/������� 15�� �����̸� ��� �޿������
'st_es_date = mid(cstr(pmg_yymm_to),1,4) & "-" & mid(cstr(pmg_yymm_to),5,2) & "-" & "01"

st_in_date = Mid(CStr(pmg_yymm_to), 1, 4)&"-"&Mid(CStr(pmg_yymm_to), 5, 2)&"-"&"16"
rever_year = Mid(CStr(pmg_yymm_to), 1, 4) '�ͼӳ⵵

'��뺸��(�Ǿ�) ����
objBuilder.Append "SELECT emp_rate, com_rate FROM pay_insurance "
objBuilder.Append "WHERE insu_yyyy = '"&rever_year&"' AND insu_id = '5503' AND insu_class = '01';"

Set rsInsEmp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsInsEmp.EOF then
	epi_emp = FormatNumber(rsInsEmp("emp_rate"), 3)
	epi_com = FormatNumber(rsInsEmp("com_rate"), 3)
Else
	epi_emp = 0
	epi_com = 0
End If
rsInsEmp.Close() : Set rsInsEmp = Nothing

'����纸�� ����
objBuilder.Append "SELECT hap_rate FROM pay_insurance "
objBuilder.Append "WHERE insu_yyyy = '"&rever_year&"' AND insu_id = '5504' AND insu_class = '01';"

Set rsInsHap = DBConn.Execute(objBuilder.ToString())
objBuilder.CleAR()

If Not rsInsHap.eof Then
	long_hap = FormatNumber(rsInsHap("hap_rate"), 3)
Else
	long_hap = 0
End If
rsInsHap.Close() : Set rsInsHap = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>�޿����� �ý���</title>
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
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">�̿����</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">�ͼӳ��</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">������</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">���</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">��  ��</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">�Ի���</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">����</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">ȸ��</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">����</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">�����</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">��</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">�μ�</div></td>
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
' �޿����޿��� 15�ϱ��� �Ի��� ����޿�ó���� ���� �޿�����Ÿ ����(�����޿������� ����)
objBuilder.Append "SELECT emtt.emp_no, emtt.emp_name, emtt.emp_company, emtt.emp_in_date, "

objBuilder.Append "	pmgt.pmg_emp_no, pmgt.pmg_grade, "
objBuilder.Append "	pmgt.pmg_company, pmgt.pmg_bonbu, pmgt.pmg_saupbu, pmgt.pmg_team, pmgt.pmg_org_name, "
objBuilder.Append "	pmgt.pmg_base_pay, pmgt.pmg_meals_pay, pmgt.pmg_postage_pay, pmgt.pmg_re_pay, "
objBuilder.Append "	pmgt.pmg_overtime_pay, pmgt.pmg_car_pay, pmgt.pmg_position_pay, pmgt.pmg_custom_pay, "
objBuilder.Append "	pmgt.pmg_job_pay, pmgt.pmg_job_support, pmgt.pmg_jisa_pay, pmgt.pmg_long_pay, "
objBuilder.Append "	pmgt.pmg_disabled_pay, pmgt.pmg_give_total, "

objBuilder.Append "	pmdt.de_emp_no, pmdt.de_nps_amt, pmdt.de_nhis_amt, pmdt.de_epi_amt, pmdt.de_longcare_amt,"
objBuilder.Append "	pmdt.de_income_tax, pmdt.de_wetax, pmdt.de_year_incom_tax, pmdt.de_year_wetax, "
objBuilder.Append "	pmdt.de_year_incom_tax2, pmdt.de_year_wetax2, pmdt.de_other_amt1, "
objBuilder.Append "	pmdt.de_sawo_amt, pmdt.de_hyubjo_amt, "
objBuilder.Append "	pmdt.de_school_amt, pmdt.de_nhis_bla_amt, pmdt.de_long_bla_amt, "

objBuilder.Append "	pyit.incom_emp_no, pyit.incom_base_pay, pyit.incom_overtime_pay, "
objBuilder.Append "	pyit.incom_month_amount, pyit.incom_family_cnt, "
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
objBuilder.Append "WHERE (ISNULL(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date >= '"&st_in_date&"') "
objBuilder.Append "	AND emtt.emp_in_date < '"&st_in_date&"' AND emtt.emp_pay_id <> '5' AND emtt.emp_no < '900000' "

If view_condi <> "��ü" Then
	objBuilder.Append "	AND emtt.emp_company = '"&view_condi&"' "
End If

objBuilder.Append "ORDER BY emtt.emp_in_date, emtt.emp_no;"

Set rsPay = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsPay.EOF Then
	arrPay = rsPay.getRows()
End If
rsPay.Close() : Set rsPay = Nothing

If IsArray(arrPay) Then
	For i = LBound(arrPay) To UBound(arrPay, 2)
		emp_no = arrPay(0, i)
		emp_name = arrPay(1, i)
		emp_company = arrPay(2, i)
		emp_in_date = arrPay(3, i)

		pmg_emp_no = f_toString(arrPay(4, i), "")
		pmg_grade = arrPay(5, i)
		pmg_company = arrPay(6, i)
		pmg_bonbu = arrPay(7, i)
		pmg_saupbu = arrPay(8, i)
		pmg_team = arrPay(9, i)
		pmg_org_name = arrPay(10, i)
		pmg_base_pay = f_toString(arrPay(11, i), 0)
		pmg_meals_pay = f_toString(arrPay(12, i), 0)
		pmg_postage_pay = f_toString(arrPay(13, i), 0)
		pmg_re_pay = f_toString(arrPay(14, i), 0)
		pmg_overtime_pay = f_toString(arrPay(15, i), 0)
		pmg_car_pay = f_toString(arrPay(16, i), 0)
		pmg_position_pay = f_toString(arrPay(17, i), 0)
		pmg_custom_pay = f_toString(arrPay(18, i), 0)
		pmg_job_pay = f_toString(arrPay(19, i), 0)
		pmg_job_support = f_toString(arrPay(20, i), 0)
		pmg_jisa_pay = f_toString(arrPay(21, i), 0)
		pmg_long_pay = f_toString(arrPay(22, i), 0)
		pmg_disabled_pay = f_toString(arrPay(23, i), 0)
		pmg_give_total = f_toString(arrPay(24, i), 0)

		de_emp_no = f_toString(arrPay(25, i), "")
		de_nps_amt = f_toString(arrPay(26, i), 0)
		de_nhis_amt = f_toString(arrPay(27, i), 0)
		de_epi_amt = f_toString(arrPay(28, i), 0)
		de_longcare_amt = f_toString(arrPay(29, i), 0)
		de_income_tax = f_toString(arrPay(30, i), 0)
		de_wetax = f_toString(arrPay(31, i), 0)
		de_year_incom_tax = f_toString(arrPay(32, i), 0)
		de_year_wetax = f_toString(arrPay(33, i), 0)
		de_year_incom_tax2 = f_toString(arrPay(34, i), 0)
		de_year_wetax2 = f_toString(arrPay(35, i), 0)
		de_other_amt1 = f_toString(arrPay(36, i), 0)
		de_sawo_amt = f_toString(arrPay(37, i), 0)
		de_hyubjo_amt = f_toString(arrPay(38, i), 0)
		de_school_amt = f_toString(arrPay(39, i), 0)
		de_nhis_bla_amt = f_toString(arrPay(40, i), 0)
		de_long_bla_amt = f_toString(arrPay(41, i), 0)

		incom_emp_no = f_toString(arrPay(42, i), "")
		incom_base_pay = f_toString(arrPay(43, i), 0)
		incom_overtime_pay = f_toString(arrPay(44, i), 0)
		incom_month_amount = f_toString(arrPay(45, i), 0)
		incom_family_cnt = f_toString(arrPay(46, i), 0)
		incom_nps = f_toString(arrPay(47, i), 0)
		incom_nhis = f_toString(arrPay(48, i), 0)
		incom_wife_yn = f_toString(arrPay(49, i), 0)
		incom_age20 = f_toString(arrPay(50, i), 0)
		incom_age60 = f_toString(arrPay(51, i), 0)
		incom_old = f_toString(arrPay(52, i), 0)
		incom_go_yn = f_toString(arrPay(53, i), "��")
		incom_long_yn = f_toString(arrPay(54, i), "��")

		If pmg_emp_no <> "" Then
			pmg_curr_pay = pmg_give_total - de_deduct_tot

		Else
			 '�⺻��/�Ĵ�� ��������
			 incom_family_cnt = 0

			 If incom_emp_no <> "" Then
				pmg_base_pay = incom_base_pay
				pmg_meals_pay = incom_meals_pay
				pmg_overtime_pay = incom_overtime_pays

				If incom_month_amount = 0 then
					  incom_month_amount = incom_base_pay + incom_overtime_pay
				Else
					  incom_month_amount = incom_month_amount
				End If
			End If

			pmg_tax_yes = pmg_base_pay + pmg_overtime_pay
			pmg_tax_no = pmg_meals_pay
			pmg_give_total = pmg_tax_yes + pmg_tax_no

			'if incom_family_cnt = 0 then
			incom_family_cnt = incom_wife_yn + incom_age20 + incom_age60 + incom_old + 1 '�ξ簡���� ������������
			'end if

			'�ٷμҵ� ���̼��� ����
			inc_st_amt = 0
			inc_incom = 0

			objBuilder.Append "SELECT inc_st_amt, inc_incom1, inc_incom2, inc_incom3, inc_incom4, inc_incom5, inc_incom6, "
			objBuilder.Append "inc_incom7, inc_incom8, inc_incom9, inc_incom10, inc_incom11 "
			objBuilder.Append "FROM pay_income_amount "
			objBuilder.Append "WHERE ('"&incom_month_amount&"' BETWEEN inc_from_amt AND inc_to_amt) AND (inc_yyyy = '"&rever_year&"');"

			Set rs_sod = DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()

			If Not rs_sod.EOF Then
				inc_st_amt = CInt(f_toString(rs_sod("inc_st_amt"), 0))

				If incom_family_cnt = 1 Then
					inc_incom = CInt(f_toString(rs_sod("inc_incom1"), 0))
				End If

				If incom_family_cnt = 2 Then
					inc_incom = CInt(f_toString(rs_sod("inc_incom2"), 0))
				End If

				If incom_family_cnt = 3 Then
					inc_incom = CInt(f_toString(rs_sod("inc_incom3"), 0))
				End If

				If incom_family_cnt = 4 Then
					inc_incom = CInt(f_toString(rs_sod("inc_incom4"), 0))
				End If

				If incom_family_cnt = 5 Then
				   inc_incom = CInt(f_toString(rs_sod("inc_incom5"), 0))
				End If

				If incom_family_cnt = 6 Then
				   inc_incom = CInt(f_toString(rs_sod("inc_incom6"), 0))
				End If

				If incom_family_cnt = 7 Then
					inc_incom = CInt(f_toString(rs_sod("inc_incom7"), 0))
				End If

				If incom_family_cnt = 8 Then
				   inc_incom = CInt(f_toString(rs_sod("inc_incom8"), 0))
				End If

				If incom_family_cnt = 9 Then
				   inc_incom = CInt(f_toString(rs_sod("inc_incom9"), 0))
				End If

				If incom_family_cnt = 10 Then
				   inc_incom = CInt(f_toString(rs_sod("inc_incom10"), 0))
				End If

				If incom_family_cnt = 11 Then
				   inc_incom = CInt(f_toString(rs_sod("inc_incom11"), 0))
				End If
			End If
			rs_sod.Close()

			'�ҵ漼
			de_income_tax = CLng(inc_incom)

			'���ο��� ���
			'nps_amt = incom_nps_amount * (nps_emp / 100)
			'nps_amt = int(nps_amt)
			'de_nps_amt = (int(nps_amt / 10)) * 10
			de_nps_amt = incom_nps

			'�ǰ����� ���
			'nhis_amt = incom_nhis_amount * (nhis_emp / 100)
			'nhis_amt = int(nhis_amt)
			'de_nhis_amt = (int(nhis_amt / 10)) * 10
			de_nhis_amt = incom_nhis

			'����纸�� ���
			If incom_long_yn = "��" Then
				long_amt = de_nhis_amt * (long_hap / 100)
				long_amt = CInt(long_amt)
				'long_amt = long_amt / 2
				de_longcare_amt = (CInt(long_amt / 10)) * 10
			Else
				de_longcare_amt = 0
			End If

			'��뺸�� ��� : ����� ������ �ݾ����� ���
			If incom_go_yn = "��" Then
				'epi_amt = inc_st_amt * (epi_emp / 100)
				epi_amt = pmg_give_tot * (epi_emp / 100)
				epi_amt = CInt(epi_amt)
				de_epi_amt = (CInt(epi_amt / 10)) * 10
			Else
				de_epi_amt = 0
			End If

			'����ҵ漼
			we_tax = inc_incom * (10 / 100)
			we_tax = CInt(we_tax)
			de_wetax = (CInt(we_tax / 10)) * 10

			de_deduct_tot = de_nps_amt + de_nhis_amt + de_epi_amt + de_longcare_amt + de_income_tax + de_wetax
			pmg_curr_pay = pmg_give_total - de_deduct_tot
		End If
%>
  <tr valign="middle" class="style11">
    <td width="110"><div align="center" class="style1"><%=pmg_yymm%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_yymm_to%></div></td>
    <td width="110"><div align="center" class="style1"><%=give_date%></div></td>
    <td width="110"><div align="center" class="style1"><%=emp_no%></div></td>
    <td width="110"><div align="center" class="style1"><%=emp_name%></div></td>
    <td width="110"><div align="center" class="style1"><%=emp_in_date%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_grade%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_company%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_bonbu%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_saupbu%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_team%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_org_name%></div></td>

    <td width="100"><div align="right" class="style1"><%=FormatNumber(pmg_base_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(pmg_meals_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(pmg_postage_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(pmg_re_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(pmg_overtime_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(pmg_car_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(pmg_position_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(pmg_custom_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(pmg_job_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(pmg_job_support, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(pmg_jisa_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(pmg_long_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(pmg_disabled_pay, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(pmg_give_total, 0)%></div></td>

    <td width="100"><div align="right" class="style1"><%=FormatNumber(de_nps_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(de_nhis_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(de_epi_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(de_longcare_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(de_income_tax, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(de_wetax, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(de_year_incom_tax, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(de_year_wetax, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(de_year_incom_tax2, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(de_year_wetax2, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(de_other_amt1, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(de_sawo_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(de_school_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(de_nhis_bla_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(de_long_bla_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(de_hyubjo_amt, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(de_deduct_tot, 0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=FormatNumber(pmg_curr_pay, 0)%></div></td>
  </tr>
<%

	Next

	Set rs_sod = Nothing
End If
DBConn.Close() : Set DBConn = Nothing
%>
</table>
</body>
</html>