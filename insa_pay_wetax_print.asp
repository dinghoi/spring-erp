<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
view_condi = request("view_condi")
pmg_yymm = request("pmg_yymm")

	sum_give_tot = 0
    sum_income_tax = 0
    sum_wetax = 0
	sum_year_incom_tax = 0
    sum_year_wetax = 0
	sum_special_tax = 0
	sum_deduct_tot = 0
	pay_count = 0	
	sum_curr_pay = 0
	
	a02_give_tot = 0
    a02_income_tax = 0
    a02_wetax = 0
	a02_count = 0	
	
	a03_give_tot = 0
    a03_income_tax = 0
    a03_wetax = 0
	a03_count = 0	
	
	a04_give_tot = 0
    a04_income_tax = 0
    a04_wetax = 0
	a04_count = 0	
	
	a10_give_tot = 0
    a10_income_tax = 0
    a10_wetax = 0
	a10_count = 0	
	
	a21_give_tot = 0
    a21_income_tax = 0
    a21_wetax = 0
	a21_count = 0	
	
	a22_give_tot = 0
    a22_income_tax = 0
    a22_wetax = 0
	a22_count = 0	
	
	a20_give_tot = 0
    a20_income_tax = 0
    a20_wetax = 0
	a20_count = 0	
	
	sum_alba_give_total = 0
    sum_tax_amt1 = 0
    sum_tax_amt2 = 0
	sum_deduct_tot = 0
	
	a32_give_tot = 0
    a32_income_tax = 0
    a32_wetax = 0
	a32_count = 0	
	
	a30_give_tot = 0
    a30_income_tax = 0
    a30_wetax = 0
	a30_count = 0
	
	tot_give_tot = 0
    tot_income_tax = 0
    tot_wetax = 0
	tot_year_incom_tax = 0
    tot_year_wetax = 0
	tot_special_tax = 0
	tot_deduct_tot = 0
	tot_pay_count = 0	
	tot_curr_pay = 0		

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = " 지방소득세계산서/납부서 "

'근로소득
Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    emp_no = rs("pmg_emp_no")
    pmg_give_tot = rs("pmg_give_total")
    pay_count = pay_count + 1
				  
    sub_give_hap = int(rs("pmg_postage_pay")) + int(rs("pmg_re_pay")) + int(rs("pmg_car_pay")) + int(rs("pmg_position_pay")) + int(rs("pmg_custom_pay")) + int(rs("pmg_job_pay")) + int(rs("pmg_job_support")) + int(rs("pmg_jisa_pay")) + int(rs("pmg_long_pay")) + int(rs("pmg_disabled_pay"))
	
	sum_give_tot = sum_give_tot + int(rs("pmg_give_total"))

    Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
    Set Rs_dct = DbConn.Execute(SQL)
    if not Rs_dct.eof then

            de_income_tax = int(Rs_dct("de_income_tax"))
            de_wetax = int(Rs_dct("de_wetax"))
			de_year_incom_tax = int(Rs_dct("de_year_incom_tax"))
            de_year_wetax = int(Rs_dct("de_year_wetax"))
		    de_deduct_tot = int(Rs_dct("de_deduct_total"))	
	     else
            de_income_tax = 0
            de_wetax = 0
			de_year_incom_tax = 0
            de_year_wetax = 0
		    de_deduct_tot = 0
     end if
     Rs_dct.close()
	 
     sum_income_tax = sum_income_tax + de_income_tax
     sum_wetax = sum_wetax + de_wetax
	 sum_year_incom_tax = sum_year_incom_tax + de_year_incom_tax
     sum_year_wetax = sum_year_wetax + de_year_wetax
	 sum_deduct_tot = sum_deduct_tot + de_deduct_tot

	rs.movenext()
loop
rs.close()

'퇴직소득

'사업소득
Sql = "select * from pay_alba_cost where (rever_yymm = '"+pmg_yymm+"' ) and (company = '"+view_condi+"') ORDER BY company,draft_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    alba_count = alba_count + 1
				  
    sum_alba_give_total = sum_alba_give_total + int(rs("alba_give_total"))
    sum_tax_amt1 = sum_tax_amt1 + int(rs("tax_amt1"))
    sum_tax_amt2 = sum_tax_amt2 + int(rs("tax_amt2"))
	sum_deduct_tot = sum_deduct_tot + (int(rs("tax_amt1")) + int(rs("tax_amt2")) + int(rs("de_other")))
	
	rs.movenext()
loop
rs.close()

'총계
tot_give_tot = a10_give_tot + a20_give_tot + a30_give_tot
tot_income_tax = sum_wetax + sum_tax_amt1
tot_wetax = a10_wetax + sum_tax_amt2
tot_pay_count = pay_count + alba_count

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>급여관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=tax_date%>" );
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
				if (document.frm.view_condi.value == "") {
					alert ("소속을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		    
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_wetax_print.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableList">
                  	<colgroup>
						<col width="30%">
						<col width="10%">
						<col width="24%">
						<col width="24%">
						<col width="12%">
					</colgroup>
				    <tbody>
                    <tr>
                      <th style="background:#FFFFE6">귀속년월</th>
                      <td colspan="2" class="left" bgcolor="#FFFFE6"><%=mid(pmg_yymm,1,4)%>년&nbsp;<%=mid(pmg_yymm,5,2)%>월분</td>
                      <th style="background:#FFFFE6">납부일</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="tax_date" type="text" size="10" id="datepicker" style="width:70px;" value="<%=tax_date%>" readonly="true">&nbsp;</td>
                    </tr>
                    <tr>
                      <th style="background:#f8f8f8;">소득구분</th>
                      <th style="background:#f8f8f8;">인원</th>
                      <th style="background:#f8f8f8;">과세표준액(소득세액)</th>
                      <th style="background:#f8f8f8;">지방소득세(소득세액의 10%)</th>
                      <th style="background:#f8f8f8;">비고</th>
                    </tr>
                    <tr>
                      <th style="background:#f8f8f8;">이 자 소 득</th>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%></td>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%>&nbsp;</td>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr>
                      <th style="background:#f8f8f8;">배 당 소 득</th>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%></td>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%>&nbsp;</td>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr>
                      <th style="background:#f8f8f8;">사 업 소 득</th>
                      <td class="right"><%=formatnumber(alba_count,0)%></td>
                      <td class="right"><%=formatnumber(sum_tax_amt1,0)%>&nbsp;</td>
                      <td class="right"><%=formatnumber(sum_tax_amt2,0)%>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr>
                      <th style="background:#f8f8f8;">갑종 근로소득</th>
                      <td class="right"><%=formatnumber(pay_count,0)%></td>
                      <td class="right"><%=formatnumber(sum_income_tax,0)%>&nbsp;</td>
                      <td class="right"><%=formatnumber(sum_wetax,0)%>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr>
                      <th style="background:#f8f8f8;">연 금 소 득</th>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%></td>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%>&nbsp;</td>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr>
                      <th style="background:#f8f8f8;">기 타 소 득</th>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%></td>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%>&nbsp;</td>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr>
                      <th style="background:#f8f8f8;">갑종 퇴직소득</th>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%></td>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%>&nbsp;</td>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr>
                      <th style="background:#f8f8f8;">외국인으로부터 받은 소득</th>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%></td>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%>&nbsp;</td>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr>
                      <th style="background:#f8f8f8;">법인세법 제98조규정에의한 특별징수</th>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%></td>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%>&nbsp;</td>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr>
                      <th style="background:#f8f8f8;">소득세법 제119조(양도소득)에의한 특별징수</th>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%></td>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%>&nbsp;</td>
                      <td class="right"><%=formatnumber(a02_income_tax,0)%>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr>
                      <th>계</th>
                      <th class="right"><%=formatnumber(tot_pay_count,0)%></th>
                      <th class="right"><%=formatnumber(tot_income_tax,0)%>&nbsp;</th>
                      <th class="right"><%=formatnumber(tot_wetax,0)%>&nbsp;</th>
                      <th>&nbsp;</th>
                    </tr>
			        </tbody>
			      </table>
				</div>
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  	 <td width="20%">
                        <div align=center>
                             <strong class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></strong>
                        </div>
				    </td>
                    <td>
					<div class="btnRight">
					<a href="#" onClick="pop_Window('insa_pay_wetax_printok.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>','insa_pay_emp_wetax_pop','scrollbars=yes,width=1060,height=700')" class="btnType04">출력</a>
					</div>                  
                    </td> 
			      </tr>
				  </table>
              </form>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="pmg_yymm" value="<%=pmg_yymm%>" ID="Hidden1">
                <input type="hidden" name="view_condi" value="<%=view_condi%>" ID="Hidden1">
		</div>				
	</body>
</html>

