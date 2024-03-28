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

main_title = CStr(curr_yyyy) & "년 " & cstr(curr_mm) & "월 " & " 급여명세서"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<title>개인 급여명세서</title>
	<script src="/java/common.js" type="text/javascript"></script>

	<script type="text/javascript">
		function printWindow(){
	//		viewOff("button");
			factory.printing.header = ""; //머리말 정의
			factory.printing.footer = ""; //꼬리말 정의
			factory.printing.portrait = true; //출력방향 설정: true - 가로, false - 세로
			factory.printing.leftMargin = 13; //외쪽 여백 설정
			factory.printing.topMargin = 25; //윗쪽 여백 설정
			factory.printing.rightMargin = 13; //오른쯕 여백 설정
			factory.printing.bottomMargin = 15; //바닦 여백 설정
	//		factory.printing.SetMarginMeasure(2); //테두리 여백 사이즈 단위를 인치로 설정
	//		factory.printing.printer = ""; //프린터 할 프린터 이름
	//		factory.printing.paperSize = "A4"; //용지선택
	//		factory.printing.pageSource = "Manusal feed"; //종이 피드 방식
	//		factory.printing.collate = true; //순서대로 출력하기
	//		factory.printing.copies = "1"; //인쇄할 매수
	//		factory.printing.SetPageRange(true,1,1); //true로 설정하고 1,3이면 1에서 3페이지 출력
	//		factory.printing.Printer(true); //출력하기
			factory.printing.Preview(); //윈도우를 통해서 출력
			factory.printing.Print(false); //윈도우를 통해서 출력
		}

		function printW() {
			window.print();
		}

		//프린트 함수 신규 작성[허정호_20220204]
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
		.style12L {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
		.style12R {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: right; }
		.style12C {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style12BC {font-size: 12px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style14BC {font-size: 14px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style14C {font-size: 14px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style14R {font-size: 14px; font-family: "바탕체", "바탕체", Seoul; text-align: right; }
		.style14L {font-size: 14px; font-family: "굴림체", "굴림체", Seoul; text-align: left; }
		.style18L {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
		.style18C {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style20L {font-size: 20px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
		.style20C {font-size: 20px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
		.style32BC {font-size: 32px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
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
		<p><a href="#" onClick="fnPrint('print_pg');"><img src="/image/printer.jpg" width="39" height="36" border="0" alt="출력하기" /></a></p>
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
		<td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14BC">사원번호</span></td>
		<td width="30%" height="30" align="left"><span class="style14C">&nbsp;&nbsp;<%=pmg_emp_no%></span></td>
		<td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14BC">사원 명</span></td>
		<td width="30%" height="30" align="left"><span class="style14C">&nbsp;&nbsp;<%=pmg_emp_name%></span></td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14BC">직 급</span></td>
		<td width="30%" height="30" align="left"><span class="style14C">&nbsp;&nbsp;<%=pmg_grade%></span></td>
		<td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14BC">입사일자</span></td>
		<td width="30%" height="30" align="left"><span class="style14C">&nbsp;&nbsp;<%=emp_in_date%></span></td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14BC">소속</span></td>
		<td width="30%" height="30" align="left"><span class="style14C">&nbsp;&nbsp;<%=pmg_org_name%>(<%=pmg_org_code%>)</span></td>
		<td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14BC">계좌번호</span></td>
		<td width="30%" height="30" align="left"><span class="style14C">&nbsp;&nbsp;<%=account_no%><br>&nbsp;&nbsp;(<%=bank_name%>-<%=account_holder%>)</span></td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center" bgcolor="#FFFFE6"><span class="style14BC">지급내역</span></td>
		<td width="30%" height="30" align="center" bgcolor="#FFFFE6"><span class="style14BC">지급액</span></td>
		<td width="20%" height="30" align="center" bgcolor="#E0FFFF"><span class="style14BC">공제내역</span></td>
		<td width="30%" height="30" align="center" bgcolor="#E0FFFF"><span class="style14BC">공제액</span></td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">기본급</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(pmg_base_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
		<td width="20%" height="30" align="center"><span class="style14C">국민연금</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(de_nps_amt, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">식대</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(pmg_meals_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">건강보험</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(de_nhis_amt, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
	  </tr>
	  <tr>
		<td width="15%" height="30" align="center"><span class="style14C">통신비</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(pmg_postage_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">고용보험</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(de_epi_amt, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">소급급여</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(pmg_re_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">장기요양보험</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(de_longcare_amt, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<tr>
		<td width="20%" height="30" align="center"><span class="style14C">연장근로수당</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(pmg_overtime_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">소득세</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(de_income_tax, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">주차지원금</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(pmg_car_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">지방소득세</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(de_wetax, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">직책수당</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(pmg_position_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
		<td width="20%" height="30" align="center"><span class="style14C">기타공제</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(de_other_amt1, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">고객관리수당</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(pmg_custom_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
		<td width="20%" height="30" align="center"><span class="style14C">사우회 회비</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(de_sawo_amt, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">직무보조비</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(pmg_job_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
		<td width="20%" height="30" align="center"><span class="style14C">협조비</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(de_hyubjo_amt, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">업무장려비</span></td>
		<td width="30%" height="30" align="right">
		<span class="style14C"><%=FormatNumber(pmg_job_support, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">학자금대출</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(de_school_amt, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">본지사근무비</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(pmg_jisa_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
		<td width="20%" height="30" align="center"><span class="style14C">건강보험료정산</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(de_nhis_bla_amt, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
	  </tr>
		<tr>
		<td width="20%" height="30" align="center"><span class="style14C">근속수당</span></td>
		<td width="30%" height="30" align="right"><span class="style14C"><%=FormatNumber(pmg_long_pay,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">장기요양보험료정산</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(de_long_bla_amt,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">장애인수당</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(pmg_disabled_pay, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
		<td width="20%" height="30" align="center"><span class="style14C">연말정산소득세</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(de_year_incom_tax, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">&nbsp;</span></td>
		<td width="30%" height="30" align="right"><span class="style14C">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">연말정산지방세</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(de_year_wetax, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">&nbsp;</span></td>
		<td width="30%" height="30" align="right"><span class="style14C">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">연말재정산소득세</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(de_year_incom_tax2, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">&nbsp;</span></td>
		<td width="30%" height="30" align="right"><span class="style14C">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center"><span class="style14C">연말재정산지방세</span></td>
		<td width="30%" height="30" align="right">
			<span class="style14C"><%=FormatNumber(de_year_wetax2, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center"><span class="style14C">&nbsp;</span></td>
		<td width="30%" height="30" align="right"><span class="style14C">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		<td width="20%" height="30" align="center" bgcolor="#E0FFFF"><span class="style14C">공제액계</span></td>
		<td width="30%" height="30" align="right" bgcolor="#E0FFFF">
			<span class="style14C"><%=FormatNumber(de_deduct_tot, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	  <tr>
		<td width="20%" height="30" align="center" bgcolor="#FFFFE6"><span class="style14C">지급액계</span></td>
		<td width="30%" height="30" align="right" bgcolor="#FFFFE6">
			<span class="style14C"><%=FormatNumber(pmg_give_tot, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
		<td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14C">차인지급액</span></td>
		<td width="30%" height="30" align="right" bgcolor="#BFBFFF">
			<span class="style14C"><%=FormatNumber(pay_curr_amt, 0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
		</td>
	  </tr>
	</table>

	<table width="690" cellpadding="0" cellspacing="0">
		<tr>
			<td width="50%" height="30" align="left" class="style1">※ 귀하의 노고에 감사드립니다</td>
			<td width="50%" height="30" align="right" valign="middle" width="100%">
			<%
			Select Case pmg_company
				Case "케이원"
					Response.Write "<img src='/image/stamp/k_one_2021_001.png' width='80' height='80' align='right'/>"
					Response.Write "<font style='font-size:14px'><br/><br/>주식회사 케이원</font>"
				Case "케이시스템"
					Response.Write "<img src='/image/stamp/k_sys_2021_001.png' width='80' height='80' align='right'/>"
					Response.Write "<font style='font-size:14px'><br/><br/>주식회사 케이시스템</font>"
				Case "케이네트웍스"
					Response.Write "<img src='/image/stamp/k_net_2021_001.png' width='80' height='80' align='right'/>"
					Response.Write "<font style='font-size:14px'><br/><br/>주식회사 케이네트웍스</font>"
				Case "에스유에이치"
					Response.Write "<img src='/image/stamp/k_one_2021_001.png' width='80' height='80' align='right'/>"
					Response.Write "<font style='font-size:14px'><br/><br/>주식회사 에스유에이치</font>"
				Case "휴디스"
					Response.Write "<img src='/image/k_hudis001.png' width='80' height='80' align='right'/>"
					Response.Write "<font style='font-size:14px'><br/><br/>주식회사 휴디스</font>"
			End Select
			%>
		<br />
		</td>
		</tr>
	</table>
	</div>
</body>
</html>