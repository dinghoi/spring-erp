<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
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
Dim sch_tab(10,10)
Dim pmg_emp_no,  pmg_emp_name, pmg_yymm, pmg_company
Dim pmg_org_code, pmg_org_name, pmg_grade
Dim pmg_base_pay, pmg_meals_pay, pmg_postage_pay, pmg_re_pay, pmg_overtime_pay
Dim pmg_car_pay, pmg_position_pay, pmg_custom_pay, pmg_job_pay, pmg_job_support
Dim pmg_jisa_pay, pmg_long_pay, pmg_disabled_pay, pmg_family_pay, pmg_school_pay
Dim pmg_qual_pay, pmg_other_pay1, pmg_other_pay2, pmg_other_pay3, pmg_tax_yes
Dim pmg_tax_no, pmg_tax_reduced, pmg_give_tot, pay_curr_atm, rsPay
Dim de_nps_amt, de_nhis_amt, de_epi_amt, de_longcare_amt, de_income_tax
Dim de_wetax, de_year_incom_tax, de_year_wetax, de_year_incom_tax2, de_year_wetax2
Dim de_other_amt1, de_sawo_amt, de_hyubjo_amt, de_school_amt, de_nhis_bla_amt
Dim de_long_bla_amt, de_deduct_tot, pay_curr_amt, main_title
Dim bank_name, account_no, account_holder, emp_in_date, curr_yyyy, curr_mm

pmg_emp_no = Request.QueryString("emp_no")
pmg_yymm = Request.QueryString("pmg_yymm")
pmg_company = Request.QueryString("pmg_company")

objBuilder.Append "SELECT pmg_emp_name, pmg_org_code, pmg_org_name, pmg_grade, pmg_base_pay, pmg_meals_pay, "
objBuilder.Append "	pmg_postage_pay, pmg_re_pay, pmg_overtime_pay, pmg_car_pay, pmg_position_pay, pmg_custom_pay, "
objBuilder.Append "	pmg_job_pay, pmg_job_support, pmg_jisa_pay, pmg_long_pay, pmg_disabled_pay, "
objBuilder.Append "	pmg_family_pay, pmg_school_pay, pmg_qual_pay, pmg_other_pay1, pmg_other_pay2, "
objBuilder.Append "	pmg_other_pay3, pmg_tax_yes, pmg_tax_no, pmg_tax_reduced, pmg_give_total, "
objBuilder.Append "	de_nps_amt, de_nhis_amt, de_epi_amt, de_longcare_amt, de_income_tax, "
objBuilder.Append "	de_wetax, de_year_incom_tax, de_year_wetax, de_year_incom_tax2, de_year_wetax2, "
objBuilder.Append "	de_other_amt1, de_special_tax, de_saving_amt, de_sawo_amt, de_johab_amt, "
objBuilder.Append "	de_hyubjo_amt, de_school_amt, de_nhis_bla_amt, de_long_bla_amt, de_deduct_total, "
objBuilder.Append "	bank_name, account_no, account_holder, emtt.emp_in_date "
objBuilder.Append "FROM pay_month_give AS pmgt "
objBuilder.Append "INNER JOIN pay_bank_account AS pbat ON pmgt.pmg_emp_no = pbat.emp_no "
objBuilder.Append "LEFT OUTER JOIN pay_month_deduct AS pmdt ON pmgt.pmg_yymm = de_yymm "
objBuilder.Append "INNER JOIN emp_master AS emtt ON pmgt.pmg_emp_no = emtt.emp_no "
objBuilder.Append "	AND pmgt.pmg_id = pmdt.de_id "
objBuilder.Append "	AND pmgt.pmg_emp_no = pmdt.de_emp_no "
objBuilder.Append "WHERE pmg_id = '1' "
objBuilder.Append "	AND pmg_yymm = '"&pmg_yymm&"' "
objBuilder.Append "	AND pmg_emp_no = '"&pmg_emp_no&"' "
objBuilder.Append "	AND pmg_company = '"&pmg_company&"';"

Set rsPay = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

pmg_emp_name = rsPay("pmg_emp_name")
pmg_org_code = rsPay("pmg_org_code")
pmg_org_name = rsPay("pmg_org_name")
pmg_grade = rsPay("pmg_grade")

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

de_nps_amt = f_toString(rsPay("de_nps_amt"), 0)
de_nhis_amt = f_toString(rsPay("de_nhis_amt"), 0)
de_epi_amt = f_toString(rsPay("de_epi_amt"), 0)
de_longcare_amt = f_toString(rsPay("de_longcare_amt"), 0)
de_income_tax = f_toString(rsPay("de_income_tax"), 0)
de_wetax = f_toString(rsPay("de_wetax"), 0)
de_year_incom_tax = f_toString(rsPay("de_year_incom_tax"), 0)
de_year_wetax = f_toString(rsPay("de_year_wetax"), 0)
de_year_incom_tax2 = f_toString(rsPay("de_year_incom_tax2"), 0)
de_year_wetax2 = f_toString(rsPay("de_year_wetax2"), 0)
de_other_amt1 = f_toString(rsPay("de_other_amt1"), 0)
de_sawo_amt = f_toString(rsPay("de_sawo_amt"), 0)
de_hyubjo_amt = f_toString(rsPay("de_hyubjo_amt"), 0)
de_school_amt = f_toString(rsPay("de_school_amt"), 0)
de_nhis_bla_amt = f_toString(rsPay("de_nhis_bla_amt"), 0)
de_long_bla_amt = f_toString(rsPay("de_long_bla_amt"), 0)
de_deduct_tot = f_toString(rsPay("de_deduct_total"), 0)

pay_curr_amt = pmg_give_tot - de_deduct_tot
emp_in_date = f_toString(rsPay("emp_in_date"), "")

bank_name = rsPay("bank_name")
account_no = rsPay("account_no")
account_holder = rsPay("account_holder")

rsPay.Close() : Set rsPay = Nothing
DBConn.Close () : Set DBConn = Nothing

curr_yyyy = Mid(CStr(pmg_yymm), 1, 4)
curr_mm = Mid(CStr(pmg_yymm), 5, 2)

main_title = CStr(curr_yyyy) & "�� " & cstr(curr_mm) & "�� " & " �޿�����"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<title>���� �޿�����</title>
	<script src="/java/common.js" type="text/javascript"></script>

	<script type="text/javascript">
		function printWindow(){
	//		viewOff("button");
			factory.printing.header = ""; //�Ӹ��� ����
			factory.printing.footer = ""; //������ ����
			factory.printing.portrait = true; //��¹��� ����: true - ����, false - ����
			factory.printing.leftMargin = 13; //���� ���� ����
			factory.printing.topMargin = 25; //���� ���� ����
			factory.printing.rightMargin = 13; //�����P ���� ����
			factory.printing.bottomMargin = 15; //�ٴ� ���� ����
	//		factory.printing.SetMarginMeasure(2); //�׵θ� ���� ������ ������ ��ġ�� ����
	//		factory.printing.printer = ""; //������ �� ������ �̸�
	//		factory.printing.paperSize = "A4"; //��������
	//		factory.printing.pageSource = "Manusal feed"; //���� �ǵ� ���
	//		factory.printing.collate = true; //������� ����ϱ�
	//		factory.printing.copies = "1"; //�μ��� �ż�
	//		factory.printing.SetPageRange(true,1,1); //true�� �����ϰ� 1,3�̸� 1���� 3������ ���
	//		factory.printing.Printer(true); //����ϱ�
			factory.printing.Preview(); //�����츦 ���ؼ� ���
			factory.printing.Print(false); //�����츦 ���ؼ� ���
		}

		function printW() {
			window.print();
		}

		//����Ʈ �Լ� �ű� �ۼ�[����ȣ_20220204]
		var printArea;
		var initBody;

		function fnPrint(id){
			printArea = document.getElementById(id);

			window.onbeforeprint = beforePrint;
			window.onafterprint = afterPrint;

			window.print();
		}

		function beforePrint(){
			initBody = document.body.innerHTML;
			document.body.innerHTML = printArea.innerHTML;
		}

		function afterPrint(){
			document.body.innerHTML = initBody;
		}
	</script>

	<style type="text/css">
	<!--
		.style12L {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: left; }
		.style12R {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: right; }
		.style12C {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style12BC {font-size: 12px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style14BC {font-size: 14px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style14C {font-size: 14px; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style14R {font-size: 14px; font-family: "����ü", "����ü", Seoul; text-align: right; }
		.style14L {font-size: 14px; font-family: "����ü", "����ü", Seoul; text-align: left; }
		.style18L {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: left; }
		.style18C {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style20L {font-size: 20px; font-family: "����ü", "����ü", Seoul; text-align: left; }
		.style20C {font-size: 20px; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style32BC {font-size: 32px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style1 {font-size:12px;color: #666666}
		.style2 {font-size:10px;color: #666666}
	-->
	</style>

	<style media="print">
		.noprint     { display: none }
	</style>
</head>
<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">

	<div class="noprint">
		<p><a href="#" onClick="fnPrint('print_pg');"><img src="/image/printer.jpg" width="39" height="36" border="0" alt="����ϱ�" /></a></p>
	</div>

	<!--<object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8"></object>-->
	<div id="print_pg">
	<table width="690" cellpadding="0" cellspacing="0">
	  <tr>
		 <td colspan="3" align="center" class="style32BC"><%=main_title%></td>
	  </tr>
	  <tr>
		 <td>&nbsp;</td>
		 <td>&nbsp;</td>
		 <td>&nbsp;</td>
	  </tr>
	</table>
	<table width="690" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000">
	  <tr>
		<td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14BC">�����ȣ</span></td>
		<td width="30%" height="30" align="left"><span class="style14C">&nbsp;&nbsp;<%=pmg_emp_no%></span></td>
		<td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14BC">��� ��</span></td>
		<td width="30%" height="30" align="left"><span class="style14C">&nbsp;&nbsp;<%=pmg_emp_name%></span></td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14BC">�� ��</span></td>
		<td width="30%" height="30" align="left"><span class="style14C">&nbsp;&nbsp;<%=pmg_grade%></span></td>
		<td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14BC">�Ի�����</span></td>
		<td width="30%" height="30" align="left"><span class="style14C">&nbsp;&nbsp;<%=emp_in_date%></span></td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14BC">�Ҽ�</span></td>
		<td width="30%" height="30" align="left"><span class="style14C">&nbsp;&nbsp;<%=pmg_org_name%>(<%=pmg_org_code%>)</span></td>
		<td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14BC">���¹�ȣ</span></td>
		<td width="30%" height="30" align="left"><span class="style14C">&nbsp;&nbsp;<%=account_no%><br>&nbsp;&nbsp;(<%=bank_name%>-<%=account_holder%>)</span></td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center" bgcolor="#FFFFE6"><span class="style14BC">���޳���</span></td>
		<td width="30%" height="30" align="center" bgcolor="#FFFFE6"><span class="style14BC">���޾�</span></td>
		<td width="20%" height="30" align="center" bgcolor="#E0FFFF"><span class="style14BC">��������</span></td>
		<td width="30%" height="30" align="center" bgcolor="#E0FFFF"><span class="style14BC">������</span></td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">�⺻��</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(pmg_base_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
		<td width="20%" height="30" align="center"><span class="style14C">���ο���</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(de_nps_amt, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">�Ĵ�</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(pmg_meals_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">�ǰ�����</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(de_nhis_amt, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
	  </tr>
	  <tr>
		<td width="15%" height="30" align="center"><span class="style14C">��ź�</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(pmg_postage_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">��뺸��</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(de_epi_amt, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">�ұޱ޿�</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(pmg_re_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">����纸��</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(de_longcare_amt, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<tr>
		<td width="20%" height="30" align="center"><span class="style14C">����ٷμ���</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(pmg_overtime_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">�ҵ漼</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(de_income_tax, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">����������</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(pmg_car_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">����ҵ漼</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(de_wetax, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">��å����</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(pmg_position_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
		<td width="20%" height="30" align="center"><span class="style14C">��Ÿ����</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(de_other_amt1, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">����������</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(pmg_custom_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
		<td width="20%" height="30" align="center"><span class="style14C">���ȸ ȸ��</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(de_sawo_amt, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">����������</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(pmg_job_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
		<td width="20%" height="30" align="center"><span class="style14C">������</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(de_hyubjo_amt, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">���������</span></td>
		<td width="30%" height="30" align="right">
		<span class="style14C"><%=FormatNumber(pmg_job_support, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">���ڱݴ���</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(de_school_amt, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">������ٹ���</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(pmg_jisa_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
		<td width="20%" height="30" align="center"><span class="style14C">�ǰ����������</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(de_nhis_bla_amt, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
	  </tr>
		<tr>
		<td width="20%" height="30" align="center"><span class="style14C">�ټӼ���</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(pmg_long_pay,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">����纸�������</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(de_long_bla_amt,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">����μ���</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(pmg_disabled_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
		<td width="20%" height="30" align="center"><span class="style14C">��������ҵ漼</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(de_year_incom_tax, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">&nbsp;</span></td>
		<td width="30%" height="30" align="right"><span class="style14C">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">�����������漼</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(de_year_wetax, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">&nbsp;</span></td>
		<td width="30%" height="30" align="right"><span class="style14C">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">����������ҵ漼</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(de_year_incom_tax2, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">&nbsp;</span></td>
		<td width="30%" height="30" align="right"><span class="style14C">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">�������������漼</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(de_year_wetax2, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">&nbsp;</span></td>
		<td width="30%" height="30" align="right"><span class="style14C">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center" bgcolor="#E0FFFF"><span class="style14C">�����װ�</span></td>
		<td width="30%" height="30" align="right" bgcolor="#E0FFFF">
			<span class="style14C"><%=FormatNumber(de_deduct_tot, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center" bgcolor="#FFFFE6"><span class="style14C">���޾װ�</span></td>
		<td width="30%" height="30" align="right" bgcolor="#FFFFE6">
			<span class="style14C"><%=FormatNumber(pmg_give_tot, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
		<td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14C">�������޾�</span></td>
		<td width="30%" height="30" align="right" bgcolor="#BFBFFF">
			<span class="style14C"><%=FormatNumber(pay_curr_amt, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	</table>

	<table width="690" cellpadding="0" cellspacing="0">
		<tr>
			<td width="50%" height="30" align="left" class="style1">�� ������ ��� ����帳�ϴ�</td>
			<td width="50%" height="30" align="right" valign="middle" width="100%">
			<%
			Select Case pmg_company
				Case "���̿�"
					Response.Write "<img src='/image/stamp/k_one_2021_001.png' width='80' height='80' align='right'/>"
					Response.Write "<font style='font-size:14px'><br/><br/>�ֽ�ȸ�� ���̿�</font>"
				Case "���̽ý���"
					Response.Write "<img src='/image/stamp/k_sys_2021_001.png' width='80' height='80' align='right'/>"
					Response.Write "<font style='font-size:14px'><br/><br/>�ֽ�ȸ�� ���̽ý���</font>"
				Case "���̳�Ʈ����"
					Response.Write "<img src='/image/stamp/k_net_2021_001.png' width='80' height='80' align='right'/>"
					Response.Write "<font style='font-size:14px'><br/><br/>�ֽ�ȸ�� ���̳�Ʈ����</font>"
				Case "����������ġ"
					Response.Write "<img src='/image/stamp/k_one_2021_001.png' width='80' height='80' align='right'/>"
					Response.Write "<font style='font-size:14px'><br/><br/>�ֽ�ȸ�� ����������ġ</font>"
				Case "�޵�"
					Response.Write "<img src='/image/k_hudis001.png' width='80' height='80' align='right'/>"
					Response.Write "<font style='font-size:14px'><br/><br/>�ֽ�ȸ�� �޵�</font>"
			End Select
			%>
		<br />
		</td>
		</tr>
	</table>
	</div>
</body>
</html>