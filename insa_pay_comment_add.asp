<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
pmg_emp_no = request("pmg_emp_no")
pmg_emp_name = request("pmg_emp_name") 
owner_view = request("owner_view")
view_company = request("view_company")
pmg_yymm = request("pmg_yymm")

pmg_comment = ""

curr_date = mid(cstr(now()),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = " 급여특이사항 등록 "
if u_type = "U" then

	'Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"') and (pmg_id = '1') and (pmg_company = '"+view_company+"') and (pmg_emp_no = '"+pmg_emp_no+"')"
	'Set rs=DbConn.Execute(Sql)

	'pmg_comment = rs("pmg_comment")
	'rs.close()

	title_line = " 급여특이사항 변경 "
	
end if

    Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"') and (pmg_id = '1') and (pmg_company = '"+view_company+"') and (pmg_emp_no = '"+pmg_emp_no+"')"
	Set rs=DbConn.Execute(Sql)

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
	
	pmg_bank_name = rs("pmg_bank_name")
	pmg_account_no = rs("pmg_account_no")
	pmg_account_holder = rs("pmg_account_holder")	
	pmg_comment = rs("pmg_comment")
	
	rs.close()
	
	Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+pmg_emp_no+"') and (de_company = '"+view_company+"')"
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

sum_curr_pay = pmg_give_tot - de_deduct_tot
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>급여관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=cmt_date%>" );
			});	  
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
				if(document.frm.cmt_comment =="") {
					alert('특이사항을 입력하세요');
					frm.cmt_comment.focus();
					return false;}
				
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_comment_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
						<col width="12%" >
						<col width="12%" >
                        <col width="12%" >
                        <col width="12%" >
						<col width="*" >
                        <col width="12%" >
                        <col width="12%" >
                        <col width="12%" >
					</colgroup>
				    <tbody>
                             <tr>
                                <th>사번</th>
                                <td class="left"><%=pmg_emp_no%>
					            <input name="pmg_emp_no" type="hidden" id="pmg_emp_no" size="9" value="<%=pmg_emp_no%>" readonly="true"></td>
                                <th>성명</th>
                                <td class="left"><%=pmg_emp_name%>
					            <input name="pmg_emp_name" type="hidden" id="pmg_emp_name" size="14" value="<%=pmg_emp_name%>" readonly="true"></td>
                                <th>직급</th>
                                <td class="left"><%=pmg_grade%>
					            <input name="pmg_grade" type="hidden" id="pmg_grade" size="9" value="<%=pmg_grade%>" readonly="true"></td>
                                <th>직책</th>
                                <td class="left"><%=pmg_position%>
					            <input name="pmg_position" type="hidden" id="pmg_position" size="14" value="<%=pmg_position%>" readonly="true"></td>
                            </tr>
                            <tr>
                                <th>소속</th>
                                <td colspan="3" class="left"><%=pmg_company%>&nbsp;&nbsp;<%=pmg_org_name%>(<%=pmg_org_code%></td>
                                <th>계좌번호</th>
                                <td colspan="3" class="left"><%=pmg_account_no%>(<%=pmg_bank_name%>-<%=pmg_account_holder%>)&nbsp;</td>
                            </tr>
					        <tr>
						        <th colspan="4" class="first" style="background:#F5FFFA">지&nbsp;급&nbsp;&nbsp;&nbsp;항&nbsp;목</th>
						        <th colspan="4" class="first" style="background:#F8F8FF">공&nbsp;제&nbsp;&nbsp;&nbsp;항&nbsp;목</th>
					        </tr>  
                            <tr>
								<th class="first" style="background:#F5FFFA">기본급</th>
								<td class="right" style="width:100px;text-align:right"><%=formatnumber(pmg_base_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">식대</th>
								<td class="right" style="width:100px;text-align:right"><%=formatnumber(pmg_meals_pay,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">국민연금</th>
                                <td class="right" style="width:100px;text-align:right"><%=formatnumber(de_nps_amt,0)%>&nbsp;</td>
                                <th style="background:#F8F8FF">건강보험</th>
                                <td class="right" style="width:100px;text-align:right"><%=formatnumber(de_nhis_amt,0)%>&nbsp;</td>
							</tr>
                        	<tr>
								<th class="first" style="background:#F5FFFA">통신비</th>
								<td class="right" style="width:100px;text-align:right"><%=formatnumber(pmg_postage_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">소급급여</th>
								<td class="right" style="width:100px;text-align:right"><%=formatnumber(pmg_re_pay,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">고용보험</th>
                                <td class="right" style="width:100px;text-align:right"><%=formatnumber(de_epi_amt,0)%>&nbsp;</td>
                                <th style="background:#F8F8FF">장기요양보험</th>
                                <td class="right" style="width:100px;text-align:right"><%=formatnumber(de_longcare_amt,0)%>&nbsp;</td>
							</tr>   
                            <tr>
								<th class="first" style="background:#F5FFFA">연장근로수당</th>
								<td class="right" style="width:100px;text-align:right"><%=formatnumber(pmg_overtime_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">주차지원금</th>
								<td class="right" style="width:100px;text-align:right"><%=formatnumber(pmg_car_pay,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">소득세</th>
                                <td class="right" style="width:100px;text-align:right"><%=formatnumber(de_income_tax,0)%>&nbsp;</td>
                                <th style="background:#F8F8FF">지방소득세</th>
                                <td class="right" style="width:100px;text-align:right"><%=formatnumber(de_wetax,0)%>&nbsp;</td>
							</tr>   
                            <tr>
								<th class="first" style="background:#F5FFFA">직책수당</th>
								<td class="right" style="width:100px;text-align:right"><%=formatnumber(pmg_position_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">고객관리수당</th>
								<td class="right" style="width:100px;text-align:right"><%=formatnumber(pmg_custom_pay,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">기타공제</th>
                                <td class="right" style="width:100px;text-align:right"><%=formatnumber(de_other_amt1,0)%>&nbsp;</td>
                                <th style="background:#F8F8FF">경조회비</th>
                                <td class="right" style="width:100px;text-align:right"><%=formatnumber(de_sawo_amt,0)%>&nbsp;</td>
							</tr>    
                            <tr>
								<th class="first" style="background:#F5FFFA">직무보조비</th>
								<td class="right" style="width:100px;text-align:right"><%=formatnumber(pmg_job_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">업무장려비</th>
								<td class="right" style="width:100px;text-align:right"><%=formatnumber(pmg_job_support,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">협조비</th>
                                <td class="right" style="width:100px;text-align:right"><%=formatnumber(de_hyubjo_amt,0)%>&nbsp;</td>
                                <th style="background:#F8F8FF">학자금대출</th>
                                <td class="right" style="width:100px;text-align:right"><%=formatnumber(de_school_amt,0)%>&nbsp;</td>
							</tr>    
                            <tr>
								<th class="first" style="background:#F5FFFA">본지사근무비</th>
								<td class="right" style="width:100px;text-align:right"><%=formatnumber(pmg_jisa_pay,0)%>&nbsp;</td>
                                <th style="background:#F5FFFA">근속수당</th>
								<td class="right" style="width:100px;text-align:right"><%=formatnumber(pmg_long_pay,0)%>&nbsp;</td>
								<th style="background:#F8F8FF">건강보험료정산</th>
                                <td class="right" style="width:100px;text-align:right"><%=formatnumber(de_nhis_bla_amt,0)%>&nbsp;</td>
                                <th style="background:#F8F8FF">장기요양보험정산</th>
                                <td class="right" style="width:100px;text-align:right"><%=formatnumber(de_long_bla_amt,0)%>&nbsp;</td>
							</tr>    
                            <tr>
								<th class="first" style=" border-bottom:2px solid #515254; background:#F5FFFA">장애인수당</th>
								<td class="right" style=" border-bottom:2px solid #515254; width:100px;text-align:right"><%=formatnumber(pmg_disabled_pay,0)%>&nbsp;</td>
                                <th style="border-bottom:2px solid #515254; background:#F5FFFA">&nbsp;</th>
								<td class="right" style=" border-bottom:2px solid #515254;">&nbsp;</td>
								<th style=" border-bottom:2px solid #515254; background:#F8F8FF">연말정산소득세</th>
                                <td class="right" style=" border-bottom:2px solid #515254; width:100px;text-align:right"><%=formatnumber(sum_year_incom_tax,0)%>&nbsp;</td>
								<th style=" border-bottom:2px solid #515254; background:#F8F8FF">연말정산<br>지방소득세</th>
                                <td class="right" style=" border-bottom:2px solid #515254; width:100px;text-align:right"><%=formatnumber(sum_year_wetax,0)%>&nbsp;</td>
							</tr>    
                            <tr>
								<th class="first" style="border-bottom:2px solid #515254; background:#F5FFFA">지급액 계</th>
								<td class="right" style=" border-bottom:2px solid #515254; width:100px;text-align:right"><%=formatnumber(pmg_give_tot,0)%>&nbsp;</td>
                                 <th style="border-bottom:2px solid #515254; background:#F5FFFA">&nbsp;</th>
								<td class="right" style=" border-bottom:2px solid #515254; width:100px;text-align:right"><%=pay_count%>&nbsp;</td>
                                <th style="border-bottom:2px solid #515254; background:#F8F8FF">공제액 계</th>
                                <td class="right" style=" border-bottom:2px solid #515254; width:100px;text-align:right"><%=formatnumber(pmg_deduct_tot,0)%>&nbsp;</td>
								<th style="border-bottom:2px solid #515254; background:#F8F8FF">차인지급액</th>
                                <td class="right" style=" border-bottom:2px solid #515254; width:100px;text-align:right"><%=formatnumber(sum_curr_pay,0)%>&nbsp;</td>
							</tr>              
                            <tr>
					            <th class="first">특이사항</th>
					            <td colspan="7" class="left">
                                <textarea name="pmg_comment" rows="1" id="textarea"><%=pmg_comment%></textarea></td>
                            </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
				<%	
				'if end_sw = "N" then	%>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
        		<%	
				'end if	%>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="pmg_company" value="<%=view_company%>" ID="Hidden1">
                <input type="hidden" name="pmg_yymm" value="<%=pmg_yymm%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

