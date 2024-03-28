<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")
pay_grade = request.cookies("nkpmg_user")("coo_pay_grade")

u_type = request("u_type")
emp_no = request("emp_no")
emp_company=Request("emp_company")

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
    sum_other_amt1 = 0
    sum_sawo_amt = 0
    sum_hyubjo_amt = 0
    sum_school_amt = 0
    sum_nhis_bla_amt = 0
    sum_long_bla_amt = 0
	sum_deduct_tot = 0
	
	pay_count = 0	
	sum_curr_pay = 0
	
curr_dd = cstr(datepart("d",now))
from_date = mid(cstr(now()-curr_dd+1),1,10)
inc_yyyy = mid(cstr(from_date),1,4)

inc_yyyyf = inc_yyyy + "01"
inc_yyyyl = inc_yyyy + "12"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_bonus = Server.CreateObject("ADODB.Recordset")
Set Rs_ins = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
Set rs_emp = DbConn.Execute(SQL)
if not rs_emp.eof then
    	emp_person1 = rs_emp("emp_person1")
		emp_person2 = rs_emp("emp_person2")
		emp_name = rs_emp("emp_name")
		emp_company = rs_emp("emp_company")
   else
		emp_person1 = ""
		emp_person2 = ""
		emp_name = ""
		emp_company = ""
end if
rs_emp.close()

Sql = "SELECT * FROM pay_year_income where incom_emp_no = '"&emp_no&"' and incom_year = '"&inc_yyyy&"'"
Set Rs_year = DbConn.Execute(SQL)
if not Rs_year.eof then
		incom_family_cnt = Rs_year("incom_family_cnt")
		incom_wife_yn = int(Rs_year("incom_wife_yn"))
		incom_age20 = Rs_year("incom_age20")
		incom_age60 = Rs_year("incom_age60")
		incom_old = Rs_year("incom_old")
   else
		incom_family_cnt = 0
		incom_wife_yn = 0
		incom_age20 = 0
		incom_age60 = 0
		incom_old = 0
end if
Rs_year.close()
bon_in = 1
incom_family_cnt = incom_wife_yn + incom_age20 + incom_age60 + incom_old + 1

title_line = "[ 소득자별근로소득원천징수부 ]"

Sql = "select * from pay_month_give where (pmg_yymm >= '"+inc_yyyyf+"' and pmg_yymm <= '"+inc_yyyyl+"') and (pmg_id = '1') and (pmg_company = '"+emp_company+"') and (pmg_emp_no = '"+emp_no+"')"
	
Rs.Open Sql, Dbconn, 1
do until rs.eof
	pmg_yymm = rs("pmg_yymm")
    pmg_give_tot = rs("pmg_give_total")
    pay_count = pay_count + 1
				  
    sum_base_pay = sum_base_pay + int(rs("pmg_base_pay"))
    sum_meals_pay = sum_meals_pay + int(rs("pmg_meals_pay"))
    sum_postage_pay = sum_postage_pay + int(rs("pmg_postage_pay"))
    sum_re_pay = sum_re_pay + int(rs("pmg_re_pay"))
    sum_overtime_pay = sum_overtime_pay + int(rs("pmg_overtime_pay"))
    sum_car_pay = sum_car_pay + int(rs("pmg_car_pay"))
    sum_position_pay = sum_position_pay + int(rs("pmg_position_pay"))
    sum_custom_pay = sum_custom_pay + int(rs("pmg_custom_pay"))
    sum_job_pay = sum_job_pay + int(rs("pmg_job_pay"))
    sum_job_support = sum_job_support + int(rs("pmg_job_support"))
    sum_jisa_pay = sum_jisa_pay + int(rs("pmg_jisa_pay"))
    sum_long_pay = sum_long_pay + int(rs("pmg_long_pay"))
    sum_disabled_pay = sum_disabled_pay + int(rs("pmg_disabled_pay"))
    sum_give_tot = sum_give_tot + int(rs("pmg_give_total"))

    Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+emp_company+"')"
    Set Rs_dct = DbConn.Execute(SQL)
    if not Rs_dct.eof then
	    	de_nps_amt = int(Rs_dct("de_nps_amt"))
            de_nhis_amt = int(Rs_dct("de_nhis_amt"))
            de_epi_amt = int(Rs_dct("de_epi_amt"))
		    de_longcare_amt = int(Rs_dct("de_longcare_amt"))
            de_income_tax = int(Rs_dct("de_income_tax"))
            de_wetax = int(Rs_dct("de_wetax"))
			de_year_incom_tax = int(Rs_dct("de_year_incom_tax"))
            de_year_wetax = int(Rs_dct("de_year_wetax"))
            de_other_amt1 = int(Rs_dct("de_other_amt1"))
            de_sawo_amt = int(Rs_dct("de_sawo_amt"))
            de_hyubjo_amt = int(Rs_dct("de_hyubjo_amt"))
            de_school_amt = int(Rs_dct("de_school_amt"))
            de_nhis_bla_amt = int(Rs_dct("de_nhis_bla_amt"))
            de_long_bla_amt = int(Rs_dct("de_long_bla_amt"))	
		    de_deduct_tot = int(Rs_dct("de_deduct_total"))	
	     else
			de_nps_amt = 0
            de_nhis_amt = 0
            de_epi_amt = 0
		    de_longcare_amt = 0
            de_income_tax = 0
            de_wetax = 0
			de_year_incom_tax = 0
            de_year_wetax = 0
            de_other_amt1 = 0
            de_sawo_amt = 0
            de_hyubjo_amt = 0
            de_school_amt = 0
            de_nhis_bla_amt = 0
            de_long_bla_amt = 0
		    de_deduct_tot = 0
     end if
     Rs_dct.close()
     sum_nps_amt = sum_nps_amt + de_nps_amt
     sum_nhis_amt = sum_nhis_amt + de_nhis_amt
     sum_epi_amt = sum_epi_amt + de_epi_amt
	 sum_longcare_amt = sum_longcare_amt + de_longcare_amt
     sum_income_tax = sum_income_tax + de_income_tax
     sum_wetax = sum_wetax + de_wetax
	 sum_year_incom_tax = sum_year_incom_tax + de_year_incom_tax
     sum_year_wetax = sum_year_wetax + de_year_wetax
     sum_other_amt1 = sum_other_amt1 + de_other_amt1
     sum_sawo_amt = sum_sawo_amt + de_sawo_amt
     sum_hyubjo_amt = sum_hyubjo_amt + de_hyubjo_amt
     sum_school_amt = sum_school_amt + de_school_amt
     sum_nhis_bla_amt = sum_nhis_bla_amt + de_nhis_bla_amt
     sum_long_bla_amt = sum_long_bla_amt + de_long_bla_amt
	 sum_deduct_tot = sum_deduct_tot + de_deduct_tot

	rs.movenext()
loop
rs.close()	
	
	
Sql = "select * from pay_month_give where (pmg_yymm >= '"+inc_yyyyf+"' and pmg_yymm <= '"+inc_yyyyl+"') and (pmg_id = '1') and (pmg_company = '"+emp_company+"') and (pmg_emp_no = '"+emp_no+"') ORDER BY pmg_emp_no,pmg_yymm"	
	
Rs.Open Sql, Dbconn, 1	

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "1 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
											$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker" ).datepicker("setDate", "<%=emp_first_date%>" );
			});
			$(function() {    $( "#datepicker1" ).datepicker();
											$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker1" ).datepicker("setDate", "<%=emp_in_date%>" );
			});	  
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function close_me(){ 
               parent.close() ;
            } 

			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			
			
			function chkfrm() {
				if(document.frm.emp_name.value =="") {
					alert('성명을 입력하세요');
					frm.emp_name.focus();
					return false;}


				a=confirm('등록하시겠습니까?'); 
				if (a==true) {
					return true;
				}
				return false;
			}
			function file_browse()	{ 
           		document.frm.att_file.click(); 
           		document.frm.text1.value=document.frm.att_file.value;  
			}
		</script>

	</head>
	<body>
    <%
    '<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onLoad="inview()">
	%>
		<div id="wrap">			

			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_emp_infor_view.asp" method="post" name="frm" enctype="multipart/form-data">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
                        <tbody>
                            <tr>
                                <td class="left">사원번호:&nbsp;<%=emp_no%>&nbsp;&nbsp;&nbsp;&nbsp;사원명:&nbsp;<%=emp_name%>&nbsp;&nbsp;&nbsp;&nbsp;주민(외국인)번호:&nbsp;<%=emp_person1%>-<%=emp_person2%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(<%=emp_company%>)</td>
 							</tr>
                            <tr>
                                <td class="left">공제대상 가족수:&nbsp;&nbsp;총인원&nbsp;&nbsp;&nbsp;&nbsp;<%=incom_family_cnt%>&nbsp;(&nbsp;본인:&nbsp;<%=bon_in%>&nbsp;&nbsp;배우자:&nbsp;<%=incom_wife_yn%>&nbsp;&nbsp;20세이하:&nbsp;<%=incom_age20%>&nbsp;&nbsp;60세이상:&nbsp;<%=incom_age60%>&nbsp;&nbsp;경로우대:&nbsp;<%=incom_old%>&nbsp;)</td>
 							</tr>
						</tbody>
					</table>
				</div>
                <br>
				<h3 class="insa" style="font-size:12px;">총급여</h3>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="10%" >
							<col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
						</colgroup>
						<thead>
							<tr>
				               <th rowspan="2" class="first" scope="col">지급년월</th>
				               <th colspan="9" scope="col" style=" border-bottom:1px solid #e3e3e3;">총&nbsp;&nbsp;&nbsp;급&nbsp;&nbsp;&nbsp;여</th>
			                </tr>
                            <tr>
								<th scope="col" style=" border-left:1px solid #e3e3e3;">급여액</th>
								<th scope="col">상여액</th>  
								<th scope="col">인정상여</th>
                                <th scope="col">주식매수 선택권<br>행사이익</th>
                                <th scope="col">우리사주 조합<br>인출금</th>
								<th scope="col">임원퇴직 소득<br>금액한도초과액</th>
                                <th scope="col">21</th>
								<th scope="col">29</th>
                                <th scope="col">계</th>
							</tr>
						</thead>                        
						<tbody>
				<% 
					 do until rs.eof
						       pmg_give_tot = rs("pmg_give_total")
							   pmg_yymm = rs("pmg_yymm")	  
       			%>                        
							<tr>
								<td class="first"><%=mid(rs("pmg_yymm"),1,4)%>년&nbsp;<%=mid(rs("pmg_yymm"),5,2)%>월</td>
                                <td class="right"><%=formatnumber(rs("pmg_give_total"),0)%></td>
                <%
						      Sql = "select * from pay_month_give where (pmg_yymm >= '"+pmg_yymm+"') and (pmg_id = '2') and (pmg_company = '"+emp_company+"') and (pmg_emp_no = '"+emp_no+"')"	
                              Set Rs_bonus = DbConn.Execute(SQL)
							  if not Rs_bonus.eof then
									bonus_give_total = int(Rs_bonus("pmg_give_total"))
	                             else
									bonus_give_total = 0
                              end if
                              Rs_bonus.close()
                %>
                                <td class="right"><%=formatnumber(bonus_give_total,0)%></td>
                <%
						      Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
                              Set Rs_dct = DbConn.Execute(SQL)
							  if not Rs_dct.eof then
									de_nps_amt = int(Rs_dct("de_nps_amt"))
                                    de_nhis_amt = int(Rs_dct("de_nhis_amt"))
                                    de_epi_amt = int(Rs_dct("de_epi_amt"))
		                            de_longcare_amt = int(Rs_dct("de_longcare_amt"))
                                    de_income_tax = int(Rs_dct("de_income_tax"))
                                    de_wetax = int(Rs_dct("de_wetax"))
									de_year_incom_tax = int(Rs_dct("de_year_incom_tax"))
                                    de_year_wetax = int(Rs_dct("de_year_wetax"))
                                    de_other_amt1 = int(Rs_dct("de_other_amt1"))
                                    de_sawo_amt = int(Rs_dct("de_sawo_amt"))
                                    de_hyubjo_amt = int(Rs_dct("de_hyubjo_amt"))
                                    de_school_amt = int(Rs_dct("de_school_amt"))
                                    de_nhis_bla_amt = int(Rs_dct("de_nhis_bla_amt"))
                                    de_long_bla_amt = int(Rs_dct("de_long_bla_amt"))	
		                            de_deduct_tot = int(Rs_dct("de_deduct_total"))	
	                             else
									de_nps_amt = 0
                                    de_nhis_amt = 0
                                    de_epi_amt = 0
		                            de_longcare_amt = 0
                                    de_income_tax = 0
                                    de_wetax = 0
									de_year_incom_tax = 0
                                    de_year_wetax = 0
                                    de_other_amt1 = 0
                                    de_sawo_amt = 0
                                    de_hyubjo_amt = 0
                                    de_school_amt = 0
                                    de_nhis_bla_amt = 0
                                    de_long_bla_amt = 0
		                            de_deduct_tot = 0
                              end if
                              Rs_dct.close()
							  pmg_curr_pay = pmg_give_tot - de_deduct_tot
							  orther_pay = 0
							  hap_pay = pmg_give_tot + bonus_give_total
							  
                %>
                                <td class="right"><%=formatnumber(orther_pay,0)%></td>
                                <td class="right"><%=formatnumber(orther_pay,0)%></td>
                                <td class="right"><%=formatnumber(orther_pay,0)%></td>
                                <td class="right"><%=formatnumber(orther_pay,0)%></td>
                                <td class="right"><%=formatnumber(orther_pay,0)%></td>
                                <td class="right"><%=formatnumber(orther_pay,0)%></td>
                                <td class="right"><%=formatnumber(hap_pay,0)%></td>
							</tr>
				<%
							rs.movenext()
						loop
						rs.close()
						
						sum_curr_pay = sum_give_tot - sum_deduct_tot
						
						sum_give_hap = sum_postage_pay + sum_re_pay + sum_car_pay + sum_position_pay + sum_custom_pay + sum_job_pay + sum_job_support + sum_jisa_pay + sum_long_pay + sum_disabled_pay
						sum_de_insur =sum_nps_amt +sum_nhis_amt +sum_epi_amt +sum_longcare_amt
						sum_de_tax =sum_income_tax +sum_wetax
						sum_de_other =sum_other_amt1 +sum_sawo_amt +sum_hyubjo_amt +sum_school_amt +sum_nhis_bla_amt +sum_long_bla_amt
						
				%>                            
						</tbody>
					</table>
				</div>
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="20%">
                        <div class="btnCenter">
                         '    <strong class="btnType01"><input type="button" value="닫기" onclick="javascript:close_me();"></strong>
                        </div> 
				    </td>
			      </tr>
				  </table>
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="view_condi" value="<%=view_condi%>" ID="Hidden1">
				</form>
		</div>				
	</div>        				
	</body>
</html>

