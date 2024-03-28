<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

u_type = request("u_type")
emp_no = request("emp_no")
emp_name = request("emp_name")
view_condi = request("view_condi")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
Set Rs_ins = Server.CreateObject("ADODB.Recordset")
Set Rs_sod = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

Sql = "SELECT * FROM emp_master where emp_no = '"+emp_no+"'"
Set rs_emp = DbConn.Execute(SQL)
if not rs_emp.eof then
    	emp_first_date = rs_emp("emp_first_date")
		emp_in_date = rs_emp("emp_in_date")
		emp_end_date = rs_emp("emp_end_date")
		emp_type = rs_emp("emp_type")
		emp_grade = rs_emp("emp_grade")
		emp_position = rs_emp("emp_position")
		emp_company = rs_emp("emp_company")
		emp_bonbu = rs_emp("emp_bonbu")
		emp_saupbu = rs_emp("emp_saupbu")
		emp_team = rs_emp("emp_team")
		emp_org_code = rs_emp("emp_org_code")
		emp_org_name = rs_emp("emp_org_name")
		emp_reside_place = rs_emp("emp_reside_place")
		emp_reside_company = rs_emp("emp_reside_company")
		emp_disabled = rs_emp("emp_disabled")
		emp_disab_grade = rs_emp("emp_disab_grade")
   else
		emp_first_date = ""
		emp_in_date = ""
		emp_end_date = ""
		emp_type = ""
		emp_grade = ""
		emp_position = ""
		emp_company = ""
		emp_bonbu = ""
		emp_saupbu = ""
		emp_team = ""
		emp_org_code = ""
		emp_org_name = ""
		emp_reside_place = ""
		emp_reside_company = ""
		emp_disabled = ""
		emp_disab_grade = ""
end if

target_date = rs_emp("emp_end_date")
emp_first_date = rs_emp("emp_first_date")
if rs_emp("emp_first_date") = "" then 
       emp_first_date = rs_emp("emp_in_date")
end if

f_year = mid(cstr(emp_first_date),1,4)
f_month = mid(cstr(emp_first_date),6,2)
f_day = mid(cstr(emp_first_date),9,2)

t_year = mid(cstr(target_date),1,4)
t_month = mid(cstr(target_date),6,2)
t_day = mid(cstr(target_date),9,2)

f_yymm = cstr(t_year) + "01"
t_yymm = cstr(t_year) + "12"

first_date = mid(cstr(target_date),1,4) + "-" + "01" + "-" + "01"

year_cnt = datediff("yyyy", first_date, target_date)
mon_cnt = datediff("m", first_date, target_date)
day_cnt = datediff("d", first_date, target_date) 

year_cnt = int(year_cnt) + 1
'mon_cnt = int(mon_cnt) + 1
mon_cnt = int(mon_cnt)
day_cnt = int(day_cnt) + 1

'response.write(year_cnt)
'response.write("/")
'response.write(mon_cnt)
'response.write("/")
'response.write(day_cnt) (pmg_yymm = '"+p_yymm+"' )

sum_give_tot = 0
sum_bunus_tot = 0
sum_tax_no = 0
sum_nps_amt = 0
sum_nhis_amt = 0
sum_epi_amt = 0
sum_longcare_amt = 0
sum_income_tax = 0
sum_wetax = 0

Sql = "select * from pay_month_give where (pmg_yymm >= '"&f_yymm&"' and pmg_yymm <= '"&t_yymm&"') and (pmg_id = '1') and (pmg_emp_no = '"+emp_no+"') and (pmg_company = '"+view_condi+"')"
Rs_give.Open Sql, Dbconn, 1
Set Rs_give = DbConn.Execute(SQL)
do until Rs_give.eof
       pmg_yymm = Rs_give("pmg_yymm")
	   pay_year = mid(cstr(Rs_give("pmg_yymm")),1,4)
            pmg_give_tot = int(Rs_give("pmg_give_total"))	
		    meals_pay = int(Rs_give("pmg_meals_pay"))	
			car_pay = int(Rs_give("pmg_car_pay"))	
	        if  meals_pay > 100000 then
			    meals_pay =  100000
	        end if
	        if  car_pay > 200000 then
			    car_pay =  200000
	        end if
	        sum_give_tot = sum_give_tot + pmg_give_tot
	        sum_tax_no = sum_tax_no + meals_pay + car_pay

  		    Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
              Set Rs_dct = DbConn.Execute(SQL)
              if not Rs_dct.eof then
                     de_nps_amt = int(Rs_dct("de_nps_amt"))	
					 de_nhis_amt = int(Rs_dct("de_nhis_amt"))	
					 de_epi_amt = int(Rs_dct("de_epi_amt"))	
					 de_longcare_amt = int(Rs_dct("de_longcare_amt"))	
					 de_income_tax = int(Rs_dct("de_income_tax"))	
					 de_wetax = int(Rs_dct("de_wetax"))	
                  else
                     de_nps_amt = 0
					 de_nhis_amt = 0
					 de_epi_amt = 0
					 de_longcare_amt = 0
					 de_income_tax = 0
					 de_wetax = 0
              end if
              Rs_dct.close()
			  sum_nps_amt = sum_nps_amt + de_nps_amt
	          sum_nhis_amt = sum_nhis_amt + de_nhis_amt
			  sum_epi_amt = sum_epi_amt + de_epi_amt
	          sum_longcare_amt = sum_longcare_amt + de_longcare_amt
			  sum_income_tax = sum_income_tax + de_income_tax
	          sum_wetax = sum_wetax + de_wetax
	Rs_give.MoveNext()
loop
Rs_give.close()
'상여금
Sql = "select * from pay_month_give where (pmg_yymm >= '"&f_yymm&"' and pmg_yymm <= '"&t_yymm&"') and (pmg_id = '2') and (pmg_emp_no = '"+emp_no+"') and (pmg_company = '"+view_condi+"')"
Rs_give.Open Sql, Dbconn, 1
Set Rs_give = DbConn.Execute(SQL)
do until Rs_give.eof
       pmg_yymm = Rs_give("pmg_yymm")
	   pay_year = mid(cstr(Rs_give("pmg_yymm")),1,4)
            pmg_give_tot = int(Rs_give("pmg_give_total"))	
	        sum_bunus_tot = sum_bunus_tot + pmg_give_tot

  		    Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '2') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
              Set Rs_dct = DbConn.Execute(SQL)
              if not Rs_dct.eof then
                     de_nps_amt = int(Rs_dct("de_nps_amt"))	
					 de_nhis_amt = int(Rs_dct("de_nhis_amt"))	
					 de_epi_amt = int(Rs_dct("de_epi_amt"))	
					 de_longcare_amt = int(Rs_dct("de_longcare_amt"))	
					 de_income_tax = int(Rs_dct("de_income_tax"))	
					 de_wetax = int(Rs_dct("de_wetax"))	
                  else
                     de_nps_amt = 0
					 de_nhis_amt = 0
					 de_epi_amt = 0
					 de_longcare_amt = 0
					 de_income_tax = 0
					 de_wetax = 0
              end if
              Rs_dct.close()
			  sum_nps_amt = sum_nps_amt + de_nps_amt
	          sum_nhis_amt = sum_nhis_amt + de_nhis_amt
			  sum_epi_amt = sum_epi_amt + de_epi_amt
	          sum_longcare_amt = sum_longcare_amt + de_longcare_amt
			  sum_income_tax = sum_income_tax + de_income_tax
	          sum_wetax = sum_wetax + de_wetax
	Rs_give.MoveNext()
loop
Rs_give.close()

Sql = "SELECT * FROM pay_year_income where incom_emp_no = '"&emp_no&"' and incom_year = '"&t_year&"'"
Set Rs_year = DbConn.Execute(SQL)
if not Rs_year.eof then
		if Rs_year("incom_month_amount") = 0 or isnull(Rs_year("incom_month_amount")) then
		        incom_month_amount = Rs_year("incom_base_pay") + Rs_year("incom_overtime_pay")
		   else
		        incom_month_amount = Rs_year("incom_month_amount")
		end if
		incom_family_cnt = Rs_year("incom_family_cnt")
		incom_nps_amount = Rs_year("incom_nps_amount")
		incom_nhis_amount = Rs_year("incom_nhis_amount")
		incom_nps = Rs_year("incom_nps")
		incom_nhis = Rs_year("incom_nhis")
		incom_wife_yn = int(Rs_year("incom_wife_yn"))
		incom_age20 = Rs_year("incom_age20")
		incom_age60 = Rs_year("incom_age60")
		incom_old = Rs_year("incom_old")
		incom_disab = Rs_year("incom_disab")
		incom_woman = int(Rs_year("incom_woman"))
    else
		incom_month_amount = 0
		incom_family_cnt = 0
		incom_nps_amount = 0
		incom_nhis_amount = 0
		incom_nps = 0
		incom_nhis = 0
		incom_wife_yn = 0
		incom_age20 = 0
		incom_age60 = 0
		incom_old = 0
		incom_disab = 0
		incom_woman = 0
end if
Rs_year.close()

total_pay = sum_give_tot + sum_bunus_tot - sum_tax_no '비과세금액을 제외한
'근로소득공제액 구하기
if total_pay >= 45000000 then
       yaer_income_amt = int(12750000 + (total_pay - 45000000) * 0.05)
   elseif total_pay >= 30000000 then
              yaer_income_amt = int(11250000 + (total_pay - 30000000) * 0.10)
		  elseif total_pay >= 15000000 then
                     yaer_income_amt = int(9000000 + (total_pay - 15000000) * 0.15)
				 elseif total_pay >= 5000000 then
                            yaer_income_amt = int(4000000 + (total_pay - 5000000) * 0.50)
						else
						    yaer_income_amt = int(total_pay * 0.70)
end if

year_soduk_amt = total_pay - yaer_income_amt '근로소득금액

'기본공제- 인적공제
bonin_amt = 1500000 '본인공제
if incom_wife_yn = 1 then '배우자공제
      wife_amt = 1500000
   else
      wife_amt = 0 
end if
family_age20 = incom_age20 * 1500000
family_age60 = incom_age60 * 1500000
family_amt = (incom_age20 + incom_age60) * 1500000 '부양가족공제
old_amt = incom_old * 1000000 '경로우대
disab_amt = incom_disab * 2000000 '장애인
if incom_woman = 1 then '부녀자공제
      woman_amt = 500000
   else
      woman_amt = 0 
end if

'연금보험료(국민연금, 연금저축등)
'sum_nps_amt 국민연금

'특별공제(보험료:건강보험,고용보험,장기요양보험 전액공제, 보장성보험등/의료비/교육비/주택자금/기부금)
total_nhis_amt = sum_nhis_amt + sum_longcare_amt '건강보험 + 장기요양보험
'sum_epi_amt 고용보험

'표준공제
sp_incom_amt = 0
'sp_incom_amt = sum_epi_amt + total_nhis_amt
'if sp_incom_amt <= 1000000 then
'       sp_incom_amt = 1000000
'   else 
'       sp_incom_amt = 0
'end if

'그 밖의 소득공제(개인연금저축/신용카드사용등.....)

'소득공제 종합한도초과액(보장성보험료+의료비+교육비+장기주택자당차입금 이자상환액+신용카등 사용금액 > 25,000,000)

'공제액계(인적공제+연금보험료공제+특별공제+그밖의소득공제+소득공제종합한도초과액)
year_deduct_hap = bonin_amt + wife_amt + family_amt + old_amt + disab_amt + woman_amt + sum_nps_amt + total_nhis_amt + sum_epi_amt

'종합소득과세표준
year_tax_sp = year_soduk_amt - year_deduct_hap

'근로소득산출세액
'if year_tax_sp >= 300000000 then
'      year_cal_tax = int(90100000 + (year_tax_sp - 300000000) * 0.38)
'   elseif year_tax_sp >= 88000000 then
'             year_cal_tax = int(15900000 + (year_tax_sp - 88000000) * 0.35)
'          elseif year_tax_sp >= 46000000 then
'                    year_cal_tax = int(5820000 + (year_tax_sp - 46000000) * 0.24)
'                 elseif year_tax_sp >= 12000000 then
'                           year_cal_tax = int(720000 + (year_tax_sp - 12000000) * 0.15)
'					    else
'						   year_cal_tax = int(year_tax_sp * 0.06)
'end if 

'근로소득산출세액(속산법:누진공제액)
if year_tax_sp >= 300000000 then
      year_cal_tax = int((year_tax_sp * 0.38) - 1940000)
   elseif year_tax_sp >= 88000000 then
             year_cal_tax = int((year_tax_sp * 0.35) - 14900000)
          elseif year_tax_sp >= 46000000 then
                    year_cal_tax = int((year_tax_sp * 0.24) - 5220000)
                 elseif year_tax_sp >= 12000000 then
                           year_cal_tax = int((year_tax_sp * 0.15) - 1080000)
					    else
						   year_cal_tax = int(year_tax_sp * 0.06)
end if 

'근로소득세액공제
if year_cal_tax >= 500000 then
       year_tax_deduct = int(275000 + (year_cal_tax - 500000) * 0.3)
   else 
       year_tax_deduct = int(year_cal_tax * 0.55)
end if
if year_tax_deduct > 500000 then
       year_tax_deduct = 500000
end if

'결정세액/지방소득세
just_income_tax = year_cal_tax - year_tax_deduct
'지방소득세
we_tax = just_income_tax * (10 / 100)
we_tax = int(we_tax)
just_wetax = (int(we_tax / 10)) * 10 

'추가징수세액
add_income_tax = just_income_tax - sum_income_tax
add_wetax = just_wetax - sum_wetax

'건강보험 요율
Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&t_year&"' and insu_id = '5502' and insu_class = '01'"
Set rs_ins = DbConn.Execute(SQL)
if not rs_ins.eof then
    	nhis_emp = formatnumber(rs_ins("emp_rate"),3)
		nhis_com = formatnumber(rs_ins("com_rate"),3)
		nhis_from = rs_ins("from_amt")
		nhis_to = rs_ins("to_amt")
   else
		nhis_emp = 0  
		nhis_com = 0
		nhis_from = 0
		his_to = 0
end if
rs_ins.close()
'장기요양보험 요율
Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&t_year&"' and insu_id = '5504' and insu_class = '01'"
Set rs_ins = DbConn.Execute(SQL)
if not rs_ins.eof then
    	long_hap = formatnumber(rs_ins("hap_rate"),3)
   else
		long_hap = 0  
end if
rs_ins.close()

re_nhis_month = int(total_pay / mon_cnt)

'건강보험 계산
nhis_amt = re_nhis_month * (nhis_emp / 100)
nhis_amt = int(nhis_amt)
re_nhis_amt = (int(nhis_amt / 10)) * 10

'장기요양보험 계산
long_amt = re_nhis_amt * (long_hap / 100)
long_amt = Int(long_amt)
'long_amt = long_amt / 2
re_longcare_amt = (Int(long_amt / 10)) * 10

re_nhis_hap = re_nhis_amt * mon_cnt
re_longcare_hap = re_longcare_amt * mon_cnt

cal_nhis_amt = re_nhis_hap - sum_nhis_amt
cal_long_amt = re_longcare_hap - sum_longcare_amt

title_line = "중도퇴직자 정산"

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
												$( "#datepicker" ).datepicker("setDate", "<%=ins_last_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=last_check_date%>" );
			});	  
			$(function() {    $( "#datepicker2" ).datepicker();
												$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker2" ).datepicker("setDate", "<%=end_date%>" );
			});	  
			$(function() {    $( "#datepicker3" ).datepicker();
												$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker3" ).datepicker("setDate", "<%=car_year%>" );
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
				if(document.frm.emp_no.value =="" ) {
					alert('사번을 입력하세요');
					frm.emp_no.focus();
					return false;}
							
				{
				a=confirm('중도퇴직자 정산처리를 하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

			function update_view() {
			var c = document.frm.u_type.value;
				if (c == 'U') 
				{
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}
        </script>
	</head>
	<body onload="update_view()">
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_empout_yearsave.asp" method="post" name="frm">
               	<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					

                        <dd>
                            <p>
                             <label>
                             <strong>사번 : </strong>
                             <input name="emp_no" type="text" value="<%=emp_no%>" style="width:50px" readonly="true">
                             -
                             <input name="emp_name" type="text" value="<%=emp_name%>" style="width:60px" readonly="true">
                             </label>
                             <label>
                             <strong>직급 : </strong>
                             <input name="emp_grade" type="text" value="<%=emp_grade%>" style="width:60px" readonly="true">
                             -
                             <input name="emp_position" type="text" value="<%=emp_position%>" style="width:70px" readonly="true">
                             </label>
                             <label>
                             <strong>입사일 : </strong>
                             <input name="emp_in_date" type="text" value="<%=emp_in_date%>" style="width:70px" readonly="true">
                             </label>
                             <label>
                             <strong>퇴직일 : </strong>
                             <input name="emp_end_date" type="text" value="<%=emp_end_date%>" style="width:70px" readonly="true">
                             </label>
                             <label>
                             <strong>소속 : </strong>
                             <input name="emp_company" type="text" value="<%=emp_company%>" style="width:90px" readonly="true">
                             -
                             <input name="emp_org_name" type="text" value="<%=emp_org_name%>" style="width:90px" readonly="true">
                             </label>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
                            <col width="*" >
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
				               <th class="first" scope="col">퇴직년도</th>
				               <th scope="col">급여</th>
                               <th scope="col">상여</th>
				               <th scope="col">비과세</th>
                               <th scope="col">소득세</th>
                               <th scope="col">지방소득세</th>
                               <th scope="col">고용보험</th>
                               <th scope="col">장기요양보험</th>
                               <th scope="col">건강보험</th>
                               <th scope="col">국민연금</th>
			               </tr>
						</thead>
						<tbody>
							<tr>
								<td class="first" style=" border-bottom:2px solid #515254;"><%=t_year%>년&nbsp;합계</td>
                                <td class="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_give_tot,0)%>&nbsp;</td>
                                <td class="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_bunus_tot,0)%>&nbsp;</td>
                                <td class="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_tax_no,0)%>&nbsp;</td>
                                <td class="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_income_tax,0)%>&nbsp;</td>
                                <td class="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_wetax,0)%>&nbsp;</td>
                                <td class="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_epi_amt,0)%>&nbsp;</td>
                                <td class="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_longcare_amt,0)%>&nbsp;</td>
                                <td class="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_nhis_amt,0)%>&nbsp;</td>
                                <td class="right" style=" border-bottom:2px solid #515254;"><%=formatnumber(sum_nps_amt,0)%>&nbsp;</td>
							</tr>
                            <tr>
								<td style="background:#f8f8f8">총급여</td>
                                <td class="right" ><%=formatnumber(total_pay,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8" >기본공제&nbsp;본인</td>
                                <td class="right" ><%=formatnumber(bonin_amt,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8" >경로우대</td>
                                <td class="right" ><%=formatnumber(old_amt,0)%>&nbsp;</td>
                                <td style="background:#f8f8f8">국민연금</td>
                                <td class="right" ><%=formatnumber(sum_nps_amt,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8">국민건강보험</td>
                                <td class="right" ><%=formatnumber(total_nhis_amt,0)%>&nbsp;</td>
						    </tr>
                            <tr>
								<td style="background:#f8f8f8">근로소득공제액</td>
                                <td class="right" ><%=formatnumber(yaer_income_amt,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8">배우자</td>
                                <td class="right" ><%=formatnumber(wife_amt,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8">장애인</td>
                                <td class="right" ><%=formatnumber(disab_amt,0)%>&nbsp;</td>
                                <td colspan="2">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8">고용보험</td>
                                <td class="right" ><%=formatnumber(sum_epi_amt,0)%>&nbsp;</td>
						    </tr>
                            <tr>
								<td style="background:#f8f8f8">근로소득금액</td>
                                <td class="right" ><%=formatnumber(year_soduk_amt,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8">부양가족</td>
                                <td class="right" ><%=formatnumber(family_amt,0)%>&nbsp;</td>
                                <td class="right" style="background:#f8f8f8">부녀자</td>
                                <td class="right" ><%=formatnumber(woman_amt,0)%>&nbsp;</td>
                                <td colspan="2">&nbsp;</td>
                                <td class="right" style="background:#f8f8f8">표준공제</td>
                                <td class="right" ><%=formatnumber(sp_incom_amt,0)%>&nbsp;</td>
						    </tr>
                            <tr>
								<td style="background:#F5FFFA">공제액계</td>
                                <td class="right" ><%=formatnumber(year_deduct_hap,0)%>&nbsp;</td>
                                <td colspan="8">&nbsp;</td>
						    </tr>
                            <tr>
								<td style="background:#F5FFFA">종합소득과세표준</td>
                                <td class="right" ><%=formatnumber(year_tax_sp,0)%>&nbsp;</td>
                                <td colspan="8">&nbsp;</td>
						    </tr>
                            <tr>
								<td style="background:#F5FFFA">근로소득산출세액</td>
                                <td class="right" ><%=formatnumber(year_cal_tax,0)%>&nbsp;</td>
                                <td style="background:#F5FFFA">결정소득세</td>
                                <td class="right" ><%=formatnumber(just_income_tax,0)%>&nbsp;</td>
                                <td style="background:#F5FFFA">기납부소득세</td>
                                <td class="right" ><%=formatnumber(sum_income_tax,0)%>&nbsp;</td>
                                <td colspan="2" class="right" style="background:#F5FFFA">추가징수 소득세</td>
                                <td class="right" ><%=formatnumber(add_income_tax,0)%>&nbsp;</td>
                                <td>&nbsp;</td>
						    </tr>
                            <tr>
								<td style="background:#F5FFFA">근로소득 세액공제</td>
                                <td class="right" ><%=formatnumber(year_tax_deduct,0)%>&nbsp;</td>
                                <td style="background:#F5FFFA">결정지방소득세</td>
                                <td class="right" ><%=formatnumber(just_wetax,0)%>&nbsp;</td>
                                <td style="background:#F5FFFA">기납부지방소득세</td>
                                <td class="right" ><%=formatnumber(sum_wetax,0)%>&nbsp;</td>
                                <td colspan="2" class="right" style="background:#F5FFFA">추가징수 지방소득세</td>
                                <td class="right" ><%=formatnumber(add_wetax,0)%>&nbsp;</td>
                                <td>&nbsp;</td>
						    </tr>
                            <tr>
								<td colspan="2" style="background:#f8f8f8">건강보험료정산</td>
                                <td colspan="8">&nbsp;</td>
						    </tr>
                            <tr>
								<td style="background:#f8f8f8">정산평균보수월액</td>
                                <td class="right" ><%=formatnumber(re_nhis_month,0)%>&nbsp;</td>
                                <td style="background:#f8f8f8">납부할&nbsp;건강보험료</td>
                                <td class="right" ><%=formatnumber(re_nhis_hap,0)%>&nbsp;</td>
                                <td style="background:#f8f8f8">납부할&nbsp;요양보험료</td>
                                <td class="right" ><%=formatnumber(re_longcare_hap,0)%>&nbsp;</td>
                                <td style="background:#f8f8f8">정산&nbsp;건강보험료</td>
                                <td class="right" ><%=formatnumber(cal_nhis_amt,0)%>&nbsp;</td>
                                <td style="background:#f8f8f8">정산&nbsp;요양보험료</td>
                                <td class="right" ><%=formatnumber(cal_long_amt,0)%>&nbsp;</td>
						    </tr>
                      </tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="중도퇴직자정산처리" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="sum_give_tot" value="<%=sum_give_tot%>" ID="Hidden1">
                <input type="hidden" name="sum_bunus_tot" value="<%=sum_bunus_tot%>" ID="Hidden1">
                <input type="hidden" name="sum_tax_no" value="<%=sum_tax_no%>" ID="Hidden1">
                <input type="hidden" name="sum_wetax" value="<%=sum_wetax%>" ID="Hidden1">
                <input type="hidden" name="sum_epi_amt" value="<%=sum_epi_amt%>" ID="Hidden1">
                <input type="hidden" name="sum_longcare_amt" value="<%=sum_longcare_amt%>" ID="Hidden1">
                <input type="hidden" name="sum_nhis_amt" value="<%=sum_nhis_amt%>" ID="Hidden1">
                <input type="hidden" name="sum_nps_amt" value="<%=sum_nps_amt%>" ID="Hidden1">
                <input type="hidden" name="total_pay" value="<%=total_pay%>" ID="Hidden1">
                <input type="hidden" name="bonin_amt" value="<%=bonin_amt%>" ID="Hidden1">
                <input type="hidden" name="total_nhis_amt" value="<%=total_nhis_amt%>" ID="Hidden1">
                <input type="hidden" name="yaer_income_amt" value="<%=yaer_income_amt%>" ID="Hidden1">
                <input type="hidden" name="wife_amt" value="<%=wife_amt%>" ID="Hidden1">
                <input type="hidden" name="year_soduk_amt" value="<%=year_soduk_amt%>" ID="Hidden1">
                <input type="hidden" name="family_amt" value="<%=family_amt%>" ID="Hidden1">
                <input type="hidden" name="sp_incom_amt" value="<%=sp_incom_amt%>" ID="Hidden1">
                <input type="hidden" name="family_age20" value="<%=family_age20%>" ID="Hidden1">
                <input type="hidden" name="family_age60" value="<%=family_age60%>" ID="Hidden1">
                <input type="hidden" name="year_deduct_hap" value="<%=year_deduct_hap%>" ID="Hidden1">
                <input type="hidden" name="year_tax_sp" value="<%=year_tax_sp%>" ID="Hidden1">
                <input type="hidden" name="year_cal_tax" value="<%=year_cal_tax%>" ID="Hidden1">
                <input type="hidden" name="just_income_tax" value="<%=just_income_tax%>" ID="Hidden1">
                <input type="hidden" name="sum_income_tax" value="<%=sum_income_tax%>" ID="Hidden1">   
                <input type="hidden" name="add_income_tax" value="<%=add_income_tax%>" ID="Hidden1">         
                <input type="hidden" name="year_tax_deduct" value="<%=year_tax_deduct%>" ID="Hidden1">
                <input type="hidden" name="just_wetax" value="<%=just_wetax%>" ID="Hidden1">
                <input type="hidden" name="add_wetax" value="<%=add_wetax%>" ID="Hidden1">     
                <input type="hidden" name="re_nhis_month" value="<%=re_nhis_month%>" ID="Hidden1">
                <input type="hidden" name="re_nhis_hap" value="<%=re_nhis_hap%>" ID="Hidden1">
                <input type="hidden" name="re_longcare_hap" value="<%=re_longcare_hap%>" ID="Hidden1">   
                <input type="hidden" name="cal_nhis_amt" value="<%=cal_nhis_amt%>" ID="Hidden1">                   
                <input type="hidden" name="cal_long_amt" value="<%=cal_long_amt%>" ID="Hidden1">                                             
			</form> 
		</div>				
	</body>
</html>

