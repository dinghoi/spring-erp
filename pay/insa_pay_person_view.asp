<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

pmg_emp_no = request("emp_no")
pmg_company = request("pmg_company")

pmg_emp_name = request("emp_name")
pmg_yymm = request("pmg_yymm")
pmg_date = request("pmg_date")
pmg_grade = request("pmg_grade")
pmg_position = request("pmg_position")

pmg_org_code = request("pmg_org_code")
pmg_org_name = request("pmg_org_name")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "월 급여지급 상세 내역"

	Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_emp_no = '"+pmg_emp_no+"') and (pmg_company = '"+pmg_company+"')"
	set rs = dbconn.execute(sql)

    pmg_yymm = rs("pmg_yymm")
	pmg_emp_no = rs("pmg_emp_no")
    pmg_company = rs("pmg_company")
	pmg_date = rs("pmg_date")
	pmg_emp_name = rs("pmg_emp_name")
	pmg_org_code = rs("pmg_org_code")
	pmg_org_name = rs("pmg_org_name")
	pmg_grade = rs("pmg_grade")
	pmg_position = rs("pmg_position")

	pmg_base_pay = rs("pmg_base_pay")
	pmg_meals_pay = rs("pmg_meals_pay")
	pmg_postage_pay = rs("pmg_postage_pay")
	pmg_re_pay = rs("pmg_re_pay")
	pmg_overtime_pay = rs("pmg_overtime_pay")
	pmg_car_pay = rs("pmg_car_pay")
	pmg_position_pay = rs("pmg_position_pay")
	pmg_custom_pay = rs("pmg_custom_pay")
	pmg_job_pay = rs("pmg_job_pay")
	pmg_job_support = rs("pmg_job_support")
	pmg_jisa_pay = rs("pmg_jisa_pay")
	pmg_long_pay = rs("pmg_long_pay")
	pmg_disabled_pay = rs("pmg_disabled_pay")
	pmg_family_pay = rs("pmg_family_pay")
	pmg_school_pay = rs("pmg_school_pay")
	pmg_qual_pay = rs("pmg_qual_pay")
	pmg_other_pay1 = rs("pmg_other_pay1")
	pmg_other_pay2 = rs("pmg_other_pay2")
	pmg_other_pay3 = rs("pmg_other_pay3")
	pmg_tax_yes = rs("pmg_tax_yes")
	pmg_tax_no = rs("pmg_tax_no")
	pmg_tax_reduced = rs("pmg_tax_reduced")
	pmg_give_tot = rs("pmg_give_total")

	rs.close()

	meals_taxno_pay = pmg_meals_pay
	car_taxno_pay = pmg_car_pay
	meals_tax_pay = 0
	car_tax_pay = 0
	if (meals_pay > 100000) then
	     meals_tax_pay = parseInt(meals_pay - 100000)
	end if
	if (meals_pay > 100000) then
	     meals_taxno_pay =  100000
	end if
	if (car_pay > 200000) then
	     car_tax_pay = parseInt(car_pay - 200000)
	end if
	if (car_pay > 200000) then
	     car_taxno_pay =  200000
	end if

	pmg_tax_yes = pmg_base_pay + pmg_postage_pay + pmg_re_pay + pmg_overtime_pay + pmg_position_pay + pmg_custom_pay + pmg_job_pay + pmg_job_support + pmg_jisa_pay + pmg_long_pay + pmg_disabled_pay + meals_tax_pay + car_tax_pay

	pmg_tax_no = meals_taxno_pay + car_taxno_pay

	Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+pmg_emp_no+"') and (de_company = '"+pmg_company+"')"
    Set Rs_dct = DbConn.Execute(SQL)
	if not Rs_dct.eof then
           de_nps_amt = Rs_dct("de_nps_amt")
           de_nhis_amt = Rs_dct("de_nhis_amt")
           de_epi_amt = Rs_dct("de_epi_amt")
		   de_longcare_amt = Rs_dct("de_longcare_amt")
           de_income_tax = Rs_dct("de_income_tax")
           de_wetax = Rs_dct("de_wetax")
		   de_year_incom_tax = Rs_dct("de_year_incom_tax")
           de_year_wetax = Rs_dct("de_year_wetax")
		   de_year_incom_tax2 = Rs_dct("de_year_incom_tax2")
           de_year_wetax2 = Rs_dct("de_year_wetax2")
           de_other_amt1 = Rs_dct("de_other_amt1")
		   if Rs_dct("de_special_tax") = "" or isnull(Rs_dct("de_special_tax")) then
		           de_special_tax = 0
		      else
			       de_special_tax = Rs_dct("de_special_tax")
		   end if
           de_saving_amt = Rs_dct("de_saving_amt")
           de_sawo_amt = Rs_dct("de_sawo_amt")
           de_johab_amt = Rs_dct("de_johab_amt")
           de_hyubjo_amt = Rs_dct("de_hyubjo_amt")
           de_school_amt = Rs_dct("de_school_amt")
           de_nhis_bla_amt = Rs_dct("de_nhis_bla_amt")
           de_long_bla_amt = Rs_dct("de_long_bla_amt")
		   de_deduct_tot = Rs_dct("de_deduct_total")
	   else
		   de_deduct_tot = 0
    end if
    Rs_dct.close()


    Sql = "SELECT * FROM pay_bank_account where emp_no = '"+pmg_emp_no+"'"
    Set rs_bnk = DbConn.Execute(SQL)
    if not rs_bnk.eof then
           bank_name = rs_bnk("bank_name")
           account_no = rs_bnk("account_no")
		   account_holder = rs_bnk("account_holder")
	   else
           bank_name = ""
		   account_no = ""
		   account_holder = ""
    end if
    rs_bnk.close()

pay_curr_amt = pmg_give_tot - de_deduct_tot

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
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
	</head>
	<% '<body onload="update_view()"> %>
    <body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_person_view.asp?emp_no=<%=emp_no%>" method="post" name="frm">
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
                                <input name="pmg_base_pay" type="text" value="<%=formatnumber(pmg_base_pay,0)%>" style="width:100px;text-align:right" readonly="true"></td>
								<th style="background:#F8F8FF">국민연금</th>
                                <td class="left">
								<input name="de_nps_amt" type="text" value="<%=formatnumber(de_nps_amt,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                        	<tr>
								<th class="first" style="background:#F5FFFA">식대</th>
								<td class="left">
                                <input name="pmg_meals_pay" type="text" value="<%=formatnumber(pmg_meals_pay,0)%>" style="width:100px;text-align:right" readonly="true"></td>
								<th style="background:#F8F8FF">건강보험</th>
                                <td class="left">
								<input name="de_nhis_amt" type="text" value="<%=formatnumber(de_nhis_amt,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                        	<tr>
								<th class="first" style="background:#F5FFFA">통신비</th>
								<td class="left">
                                <input name="pmg_postage_pay" type="text" value="<%=formatnumber(pmg_postage_pay,0)%>" style="width:100px;text-align:right" readonly="true"></td>
								<th style="background:#F8F8FF">고용보험</th>
                                <td class="left">
								<input name="de_epi_amt" type="text" value="<%=formatnumber(de_epi_amt,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">소급급여</th>
								<td class="left">
                                <input name="pmg_re_pay" type="text" value="<%=formatnumber(pmg_re_pay,0)%>" style="width:100px;text-align:right" readonly="true"></td>
								<th style="background:#F8F8FF">장기요양보험</th>
                                <td class="left">
								<input name="de_longcare_amt" type="text" value="<%=formatnumber(de_longcare_amt,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">연장근로수당</th>
								<td class="left">
                                <input name="pmg_overtime_pay" type="text" value="<%=formatnumber(pmg_overtime_pay,0)%>" style="width:100px;text-align:right" readonly="true"></td>
								<th style="background:#F8F8FF">소득세</th>
                                <td class="left">
								<input name="de_income_tax" type="text" value="<%=formatnumber(de_income_tax,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">주차지원금</th>
								<td class="left">
                                <input name="pmg_car_pay" type="text" value="<%=formatnumber(pmg_car_pay,0)%>" style="width:100px;text-align:right" readonly="true"></td>
								<th style="background:#F8F8FF">지방소득세</th>
                                <td class="left">
								<input name="de_wetax" type="text" value="<%=formatnumber(de_wetax,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">직책수당</th>
								<td class="left">
                                <input name="pmg_position_pay" type="text" value="<%=formatnumber(pmg_position_pay,0)%>" style="width:100px;text-align:right" readonly="true"></td>
								<th style="background:#F8F8FF">기타공제</th>
                                <td class="left">
								<input name="de_other_amt1" type="text" value="<%=formatnumber(de_other_amt1,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">고객관리수당</th>
								<td class="left">
                                <input name="pmg_custom_pay" type="text" value="<%=formatnumber(pmg_custom_pay,0)%>" style="width:100px;text-align:right" readonly="true"></td>
								<th style="background:#F8F8FF">경조회비</th>
                                <td class="left">
								<input name="de_sawo_amt" type="text" value="<%=formatnumber(de_sawo_amt,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">직무보조비</th>
								<td class="left">
                                <input name="pmg_job_pay" type="text" value="<%=formatnumber(pmg_job_pay,0)%>" style="width:100px;text-align:right" readonly="true"></td>
								<th style="background:#F8F8FF">협조비</th>
                                <td class="left">
								<input name="de_hyubjo_amt" type="text" value="<%=formatnumber(de_hyubjo_amt,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">업무장려비</th>
								<td class="left">
                                <input name="pmg_job_support" type="text" value="<%=formatnumber(pmg_job_support,0)%>" style="width:100px;text-align:right" readonly="true"></td>
								<th style="background:#F8F8FF">학자금대출</th>
                                <td class="left">
								<input name="de_school_amt" type="text" value="<%=formatnumber(de_school_amt,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">본지사근무비</th>
								<td class="left">
                                <input name="pmg_jisa_pay" type="text" value="<%=formatnumber(pmg_jisa_pay,0)%>" style="width:100px;text-align:right" readonly="true"></td>
								<th style="background:#F8F8FF">건강보험료정산</th>
                                <td class="left">
								<input name="de_nhis_bla_amt" type="text" value="<%=formatnumber(de_nhis_bla_amt,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">근속수당</th>
								<td class="left">
                                <input name="pmg_long_pay" type="text" value="<%=formatnumber(pmg_long_pay,0)%>" style="width:100px;text-align:right" readonly="true"></td>
								<th style="background:#F8F8FF">장기요양보험정산</th>
                                <td class="left">
								<input name="de_long_bla_amt" type="text" value="<%=formatnumber(de_long_bla_amt,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                            <tr>
								<th class="first" style=" border-bottom:1px solid #e3e3e3; background:#F5FFFA">장애인수당</th>
								<td class="left">
                                <input name="pmg_disabled_pay" type="text" value="<%=formatnumber(pmg_disabled_pay,0)%>" style="width:100px;text-align:right" readonly="true"></td>
								<th style="background:#F8F8FF">연말정산소득세</th>
                                <td class="left">
								<input name="de_year_incom_tax" type="text" value="<%=formatnumber(de_year_incom_tax,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA"></th>
								<td class="left">&nbsp;</td>
                                <input name="pmg_family_pay" type="hidden" value="<%=formatnumber(pmg_family_pay,0)%>" style="width:100px;text-align:right" onKeyUp="give_cal(this);"></td>
								<th style="background:#F8F8FF">연말정산지방세</th>
                                <td class="left">
								<input name="de_year_wetax" type="text" value="<%=formatnumber(de_year_wetax,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">과세</th>
								<td class="left">
                                <input name="pmg_tax_yes" type="text" value="<%=formatnumber(pmg_tax_yes,0)%>" style="width:100px;text-align:right" readonly="true"></td>
								<th style="background:#F8F8FF">연말재정산소득세</th>
                                <td class="left">
								<input name="de_year_incom_tax2" type="text" value="<%=formatnumber(de_year_incom_tax2,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">비과세</th>
								<td class="left">
                                <input name="pmg_tax_no" type="text" value="<%=formatnumber(pmg_tax_no,0)%>" style="width:100px;text-align:right" readonly="true"></td>
								<th style="background:#F8F8FF">연말재정산지방세</th>
                                <td class="left">
								<input name="de_year_wetax2" type="text" value="<%=formatnumber(de_year_wetax2,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">감면소득</th>
								<td class="left">
                                <input name="pmg_tax_reduced" type="text" value="<%=formatnumber(pmg_tax_reduced,0)%>" style="width:100px;text-align:right" readonly="true"></td>
								<th style="background:#F8F8FF">공제액 계</th>
                                <td class="left">
								<input name="de_deduct_tot" type="text" value="<%=formatnumber(de_deduct_tot,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">지급액 계</th>
								<td class="left">
                                <input name="pmg_give_tot" type="text" value="<%=formatnumber(pmg_give_tot,0)%>" style="width:100px;text-align:right" readonly="true"></td>
								<th style="background:#F8F8FF">차인지급액</th>
                                <td class="left">
								<input name="pay_curr_amt" type="text" value="<%=formatnumber(pay_curr_amt,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                      </tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01">
                    <a href="#" onClick="pop_Window('insa_pay_person_print.asp?emp_no=<%=pmg_emp_no%>&emp_name=<%=pmg_emp_name%>&pmg_yymm=<%=pmg_yymm%>&pmg_date=<%=pmg_date%>&pmg_company=<%=pmg_company%>&pmg_org_code=<%=pmg_org_code%>&pmg_org_name=<%=pmg_org_name%>&pmg_grade=<%=pmg_grade%>&pmg_position=<%=pmg_position%>','insa_pop_report','scrollbars=yes,width=750,height=700')"><input type="button" value="출력" ID="Button1" NAME="Button1"></a>
			        </span>
                    <span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>
                </div>
			</form>
		</div>
	</body>
</html>
