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
Dim pmg_emp_no, pmg_yymm, title_line, rsPay
Dim pmg_company, pmg_date, pmg_emp_name, pmg_org_name, pmg_grade
Dim pmg_org_code, pmg_position, pmg_base_pay, pmg_meals_pay, pmg_postage_pay
Dim pmg_re_pay, pmg_overtime_pay, pmg_car_pay, pmg_position_pay, pmg_custom_pay
Dim pmg_job_pay, pmg_job_support, pmg_jisa_pay, pmg_long_pay, pmg_disabled_pay
Dim pmg_family_pay, pmg_school_pay, pmg_qual_pay, pmg_other_pay1, pmg_other_pay2
Dim pmg_other_pay3, pmg_tax_yes, pmg_tax_no, pmg_tax_reduced, pmg_give_tot
Dim de_nps_amt, de_nhis_amt, de_epi_amt, de_longcare_amt, de_income_tax
Dim de_wetax, de_year_incom_tax, de_year_wetax, de_year_incom_tax2, de_year_wetax2
Dim de_other_amt1, de_special_tax, de_saving_amt, de_sawo_amt, de_johab_amt
Dim de_hyubjo_amt, de_school_amt, de_nhis_bla_amt, de_long_bla_amt, de_deduct_tot
Dim meals_taxno_pay, car_taxno_pay, meals_tax_pay, car_tax_pay, meals_pay
Dim pmg_tax_pay_1, pmg_tax_pay_2, pmg_tax_pay_3
Dim bank_name, account_no, account_holder
Dim pay_curr_amt

pmg_emp_no = Request.QueryString("emp_no")
pmg_yymm = Request.QueryString("pmg_yymm")
pmg_company = Request.QueryString("pmg_company")

title_line = Left(pmg_yymm, 4)&"년 "&Mid(pmg_yymm, 5, 2)&"월 급여지급 상세 내역"

objBuilder.Append "SELECT pmg_company, pmg_date, pmg_emp_name, pmg_org_code, pmg_org_name, "
objBuilder.Append "	pmg_grade, pmg_position, pmg_base_pay, pmg_meals_pay, pmg_postage_pay, "
objBuilder.Append "	pmg_re_pay, pmg_overtime_pay, pmg_car_pay, pmg_position_pay, pmg_custom_pay, "
objBuilder.Append "	pmg_job_pay, pmg_job_support, pmg_jisa_pay, pmg_long_pay, pmg_disabled_pay, "
objBuilder.Append "	pmg_family_pay, pmg_school_pay, pmg_qual_pay, pmg_other_pay1, pmg_other_pay2, "
objBuilder.Append "	pmg_other_pay3, pmg_tax_yes, pmg_tax_no, pmg_tax_reduced, pmg_give_total, "
objBuilder.Append "	de_nps_amt, de_nhis_amt, de_epi_amt, de_longcare_amt, de_income_tax, "
objBuilder.Append "	de_wetax, de_year_incom_tax, de_year_wetax, de_year_incom_tax2, de_year_wetax2, "
objBuilder.Append "	de_other_amt1, de_special_tax, de_saving_amt, de_sawo_amt, de_johab_amt, "
objBuilder.Append "	de_hyubjo_amt, de_school_amt, de_nhis_bla_amt, de_long_bla_amt, de_deduct_total, "
objBuilder.Append "	bank_name, account_no, account_holder "
objBuilder.Append "FROM pay_month_give AS pmgt "
objBuilder.Append "INNER JOIN pay_bank_account AS pbat ON pmgt.pmg_emp_no = pbat.emp_no "
objBuilder.Append "LEFT OUTER JOIN pay_month_deduct AS pmdt ON pmgt.pmg_yymm = de_yymm "
objBuilder.Append "	AND pmgt.pmg_id = pmdt.de_id "
objBuilder.Append "	AND pmgt.pmg_emp_no = pmdt.de_emp_no "
objBuilder.Append "	AND pmgt.pmg_company = pmdt.de_company "
objBuilder.Append "WHERE pmg_id = '1' "
objBuilder.Append "	AND pmg_yymm = '"&pmg_yymm&"' "
objBuilder.Append "	AND pmg_emp_no = '"&pmg_emp_no&"' "
objBuilder.Append "	AND pmg_company = '"&pmg_company&"';"

Set rsPay = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

pmg_company = rsPay("pmg_company")
pmg_date = rsPay("pmg_date")
pmg_emp_name = rsPay("pmg_emp_name")
pmg_org_code = rsPay("pmg_org_code")
pmg_org_name = rsPay("pmg_org_name")
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
de_special_tax = f_toString(rsPay("de_special_tax"), 0)
de_saving_amt = f_toString(rsPay("de_saving_amt"), 0)
de_sawo_amt = f_toString(rsPay("de_sawo_amt"), 0)
de_johab_amt = f_toString(rsPay("de_johab_amt"), 0)
de_hyubjo_amt = f_toString(rsPay("de_hyubjo_amt"), 0)
de_school_amt = f_toString(rsPay("de_school_amt"), 0)
de_nhis_bla_amt = f_toString(rsPay("de_nhis_bla_amt"), 0)
de_long_bla_amt = f_toString(rsPay("de_long_bla_amt"), 0)
de_deduct_tot = f_toString(rsPay("de_deduct_total"), 0)

bank_name = f_toString(rsPay("bank_name"), "")
account_no = f_toString(rsPay("account_no"), "")
account_holder = f_toString(rsPay("account_holder"), "")

rsPay.Close() : Set rsPay = Nothing
DBConn.Close() : Set DBConn = Nothing

meals_taxno_pay = pmg_meals_pay
car_taxno_pay = pmg_car_pay
meals_tax_pay = 0
car_tax_pay = 0

'비정상적인 코드 확인으로 우선 주석처리[허정호_20210723]
'-> meals_pay, car_pay 변수는 해당 코드 내 정의되지 않음(필요없는 조건 코드)
'if (meals_pay > 100000) then
'	 meals_tax_pay = parseInt(meals_pay - 100000)
'end If
'if (meals_pay > 100000) then
'	 meals_taxno_pay =  100000
'end If
'if (car_pay > 200000) then
'	 car_tax_pay = parseInt(car_pay - 200000)
'end if
'if (car_pay > 200000) then
'	 car_taxno_pay =  200000
'end if

pmg_tax_pay_1 = pmg_base_pay + pmg_postage_pay + pmg_re_pay + pmg_overtime_pay
pmg_tax_pay_2 = pmg_position_pay + pmg_custom_pay + pmg_job_pay + pmg_job_support
pmg_tax_pay_3 = pmg_jisa_pay + pmg_long_pay + pmg_disabled_pay + meals_tax_pay + car_tax_pay

pmg_tax_yes = pmg_tax_pay_1 + pmg_tax_pay_2 + pmg_tax_pay_3
pmg_tax_no = meals_taxno_pay + car_taxno_pay

pay_curr_amt = pmg_give_tot - de_deduct_tot
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			function chkfrm() {
				if(document.frm.emp_no.value =="") {
					alert('성명을 입력하세요');
					frm.emp_no.focus();
					return false;}
				{
					return true;
				}
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
    <body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="container">
			<h3 class="insa"><%=title_line%></h3><br/>
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
							<td class="left"><%=pmg_emp_no%>&nbsp;</td>
							<th >성명</th>
							<td class="left" ><%=pmg_emp_name%>&nbsp;</td>
						</tr>
						<tr>
							<th class="first">직급</th>
							<td class="left"><%=pmg_grade%>&nbsp;</td>
							<th >직책</th>
							<td class="left" ><%=pmg_position%>&nbsp;</td>
						</tr>
						<tr>
							<th class="first">귀속년월</th>
							<td class="left" ><%=pmg_yymm%>&nbsp;</td>
							<th >지급일</th>
							<td class="left"><%=pmg_date%>&nbsp;</td>
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
							<th class="first" style="background:#F5FFFA">기본급</th>
							<td class="left">
								<input type="text" name="pmg_base_pay" value="<%=FormatNumber(pmg_base_pay, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">국민연금</th>
							<td class="left">
								<input type="text" name="de_nps_amt" value="<%=FormatNumber(de_nps_amt, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA">식대</th>
							<td class="left">
							<input type="text" name="pmg_meals_pay" value="<%=FormatNumber(pmg_meals_pay, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/></td>
							<th style="background:#F8F8FF">건강보험</th>
							<td class="left">
							<input type="text" name="de_nhis_amt" value="<%=FormatNumber(de_nhis_amt, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/></td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA">통신비</th>
							<td class="left">
								<input type="text" name="pmg_postage_pay" value="<%=FormatNumber(pmg_postage_pay, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">고용보험</th>
							<td class="left">
								<input type="text" name="de_epi_amt" value="<%=FormatNumber(de_epi_amt, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA">소급급여</th>
							<td class="left">
								<input type="text" name="pmg_re_pay" value="<%=FormatNumber(pmg_re_pay, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">장기요양보험</th>
							<td class="left">
								<input type="text" name="de_longcare_amt" value="<%=FormatNumber(de_longcare_amt, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA">연장근로수당</th>
							<td class="left">
								<input type="text" name="pmg_overtime_pay" value="<%=formatnumber(pmg_overtime_pay,0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">소득세</th>
							<td class="left">
								<input type="text" name="de_income_tax" value="<%=FormatNumber(de_income_tax, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA">주차지원금</th>
							<td class="left">
								<input type="text" name="pmg_car_pay" value="<%=FormatNumber(pmg_car_pay, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">지방소득세</th>
							<td class="left">
								<input type="text" name="de_wetax" value="<%=FormatNumber(de_wetax, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA">직책수당</th>
							<td class="left">
								<input type="text" name="pmg_position_pay" value="<%=FormatNumber(pmg_position_pay, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">기타공제</th>
							<td class="left">
								<input type="text" name="de_other_amt1" value="<%=FormatNumber(de_other_amt1, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA">고객관리수당</th>
							<td class="left">
								<input type="text" name="pmg_custom_pay" value="<%=FormatNumber(pmg_custom_pay, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">경조회비</th>
							<td class="left">
								<input type="text" name="de_sawo_amt" value="<%=FormatNumber(de_sawo_amt, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA">직무보조비</th>
							<td class="left">
								<input type="text" name="pmg_job_pay" value="<%=FormatNumber(pmg_job_pay, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">협조비</th>
							<td class="left">
								<input type="text" name="de_hyubjo_amt" value="<%=FormatNumber(de_hyubjo_amt, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA">업무장려비</th>
							<td class="left">
								<input type="text" name="pmg_job_support" value="<%=FormatNumber(pmg_job_support, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">학자금대출</th>
							<td class="left">
								<input type="text" name="de_school_amt" value="<%=FormatNumber(de_school_amt, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA">본지사근무비</th>
							<td class="left">
								<input type="text" name="pmg_jisa_pay" value="<%=FormatNumber(pmg_jisa_pay, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">건강보험료정산</th>
							<td class="left">
								<input type="text" name="de_nhis_bla_amt" value="<%=FormatNumber(de_nhis_bla_amt, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA">근속수당</th>
							<td class="left">
								<input type="text" name="pmg_long_pay" value="<%=FormatNumber(pmg_long_pay, 0)%>" style="width:100px;text-align:right" class="no-input"readonly/>
							</td>
							<th style="background:#F8F8FF">장기요양보험정산</th>
							<td class="left">
								<input type="text" name="de_long_bla_amt" value="<%=FormatNumber(de_long_bla_amt, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA">장애인수당</th>
							<td class="left">
								<input type="text" name="pmg_disabled_pay" value="<%=FormatNumber(pmg_disabled_pay, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">연말정산소득세</th>
							<td class="left">
								<input type="text" name="de_year_incom_tax" value="<%=FormatNumber(de_year_incom_tax, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA">가족수당</th>
							<td class="left">
								 <input type="text" name="pmg_family_pay" value="<%=FormatNumber(pmg_family_pay, 0)%>" style="width:100px;text-align:right;" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">연말정산지방세</th>
							<td class="left">
								<input type="text" name="de_year_wetax" value="<%=FormatNumber(de_year_wetax, 0)%>" style="width:100px;text-align:right;" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA">과세</th>
							<td class="left">
								<input type="text" name="pmg_tax_yes" value="<%=FormatNumber(pmg_tax_yes, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">연말재정산소득세</th>
							<td class="left">
								<input type="text" name="de_year_incom_tax2" value="<%=FormatNumber(de_year_incom_tax2, 0)%>" style="width:100px;text-align:right" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA">비과세</th>
							<td class="left">
								<input type="text" name="pmg_tax_no" value="<%=FormatNumber(pmg_tax_no, 0)%>" style="width:100px;text-align:right;" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">연말재정산지방세</th>
							<td class="left">
								<input type="text" name="de_year_wetax2" value="<%=FormatNumber(de_year_wetax2, 0)%>" style="width:100px;text-align:right;" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA">감면소득</th>
							<td class="left">
								<input type="text" name="pmg_tax_reduced" value="<%=FormatNumber(pmg_tax_reduced, 0)%>" style="width:100px;text-align:right;" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">공제액 계</th>
							<td class="left">
								<input type="text" name="de_deduct_tot" value="<%=FormatNumber(de_deduct_tot, 0)%>" style="width:100px;text-align:right;" class="no-input" readonly/>
							</td>
						</tr>
						<tr>
							<th class="first" style="background:#F5FFFA">지급액 계</th>
							<td class="left">
								<input type="text" name="pmg_give_tot" value="<%=FormatNumber(pmg_give_tot, 0)%>" style="width:100px;text-align:right;" class="no-input" readonly/>
							</td>
							<th style="background:#F8F8FF">차인지급액</th>
							<td class="left">
								<input type="text" name="pay_curr_amt" value="<%=FormatNumber(pay_curr_amt, 0)%>" style="width:100px;text-align:right;" class="no-input" readonly/>
							</td>
						</tr>
				  </tbody>
				</table>
			</div>
			<br>
			<div align="center">
				<span class="btnType01">
					<a href="#" onClick="pop_Window('/person/insa_pay_person_print.asp?emp_no=<%=pmg_emp_no%>&pmg_yymm=<%=pmg_yymm%>&pmg_company=<%=pmg_company%>','급여 출력','scrollbars=yes,width=720,height=700')"><input type="button" value="출력"/></a>
				</span>
				<span class="btnType01"><input type="button" value="닫기" onclick="toclose();"/></span>
			</div>
		</div>
	</body>
</html>