<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim pay_tab(50,10)
dim pay_pay(50,10)
dim bonus_tab(50,10)

u_type = request("u_type")
emp_no = request("emp_no")
emp_name = request("emp_name")
eot_empno = request("emp_no")
eot_emp_name = request("emp_name")
view_condi = request("view_condi")
eot_end_date = request("eot_end_date")
eot_st_date = request("eot_st_date")
eot_cen_date = request("eot_cen_date")

	eot_emp_name = ""
	eot_in_date = ""
	eot_end_date = ""
	eot_company = ""
	eot_bonbu = ""
	eot_saupbu = ""
	eot_team = ""
	eot_org_code = ""
	eot_org_name = ""
	eot_comment = ""
	eot_space = ""
	eot_mon1 = 0
	eot_mon2 = 0
	eot_mon3 = 0
	eot_work_cnt = 0
	eot_bonus = 0
	eot_mon_hap = 0
	eot_bonus_hap = 0
	eot_pay_sum = 0
	eot_average_pay = 0
	eot_end_pay = 0
	
	eos_emp_name = ""
	eos_first_date = ""
	eos_in_date = ""
	eos_end_date = ""
	eos_company = ""
	eos_bonbu = ""
	eos_saupbu = ""
	eos_team = ""
	eos_org_code = ""
	eos_org_name = ""
	eos_sum_end_pay = 0
	eos_incomm_tax = 0
	eos_we_tax = 0
	eos_tax_hap = 0
	eos_loan_jan = 0
	eos_loan_interest = 0
	eos_loan_hap = 0
	eos_nhis_amt = 0
	eos_give_amt = 0
	eos_end_pension = 0
	eos_curr_give_amt = 0

for i = 1 to 50
    for j = 1 to 10
	    pay_tab(i,j) = ""
		pay_pay(i,j) = 0
		bonus_tab(i,j) = 0
    next
next


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
Set Rs_ytax = Server.CreateObject("ADODB.Recordset")
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
end_year = mid(cstr(target_date),1,4)

emp_first_date = rs_emp("emp_first_date")
if rs_emp("emp_first_date") = "" then 
       emp_first_date = rs_emp("emp_in_date")
end if
'target_date = "2015-02-20"
'emp_first_date = "2013-11-10"

f_year = int(mid(cstr(emp_first_date),1,4))
f_month = int(mid(cstr(emp_first_date),6,2))
f_day = int(mid(cstr(emp_first_date),9,2))
fcal_day = cstr(f_day)
cf_date = emp_first_date '중간퇴직처리를 하기위한

t_year = int(mid(cstr(target_date),1,4))
t_month = int(mid(cstr(target_date),6,2))
t_day = int(mid(cstr(target_date),9,2))
tcal_month = mid(cstr(target_date),1,4) + mid(cstr(target_date),6,2)
tcal_day = cstr(t_day)

year_cnt = datediff("yyyy", emp_first_date, target_date)
mon_cnt = datediff("m", emp_first_date, target_date)
day_cnt = datediff("d", emp_first_date, target_date) 

year_cnt = int(year_cnt) + 1
mon_cnt = int(mon_cnt) + 1
day_cnt = int(day_cnt) + 1

'response.write(year_cnt)
'response.write("/")
'response.write(mon_cnt)
'response.write("/")
'response.write(day_cnt)

i = 0
j = 0
if mon_cnt >= 12 then ' 1월입사 12월퇴사인경우
   if f_year = t_year then
      year1_cnt = int(datediff("d", emp_first_date, target_date)) + 1
	  pay_tab(i,6) = year1_cnt
	  start_date = cstr(mid(emp_first_date,1,4) + "-" + "01" + "-" + "01")
	  end_date = cstr(mid(emp_first_date,1,4) + "-" + "12" + "-" + "31")
	  pay_tab(i,7) = int(datediff("d", start_date, end_date)) + 1
	  if t_day >= 15 then
            i = i + 1
	        pay_tab(i,1) = cstr(f_year)
			pay_tab(i,4) = cstr(tcal_month)
	        tcal_month = cstr(int(tcal_month) - 1)
	        pay_tab(i,3) = cstr(tcal_month)
	        tcal_month = cstr(int(tcal_month) - 1)
	        pay_tab(i,2) = cstr(tcal_month)
			tar1_date = cstr(mid(pay_tab(i,4),1,4) + "-" + mid(pay_tab(i,4),5,2) + "-" + tcal_day)
			fir1_date = cstr(mid(pay_tab(i,2),1,4) + "-" + mid(pay_tab(i,2),5,2) + "-" + "01")
			work1_cnt = int(datediff("d", fir1_date, tar1_date)) + 1
			pay_tab(i,5) = work1_cnt
	      else 
		    i = i + 1
			pay_tab(i,1) = cstr(f_year)
			tcal_month = cstr(int(tcal_month) - 1)
	        pay_tab(i,4) = cstr(tcal_month)
	        tcal_month = cstr(int(tcal_month) - 1)
	        pay_tab(i,3) = cstr(tcal_month)
	        tcal_month = cstr(int(tcal_month) - 1)
	        pay_tab(i,2) = cstr(tcal_month)
			tar1_date = cstr(mid(pay_tab(i,4),1,4) + "-" + mid(pay_tab(i,4),5,2) + "-" + tcal_day)
			fir1_date = cstr(mid(pay_tab(i,2),1,4) + "-" + mid(pay_tab(i,2),5,2) + "-" + "01")
			work1_cnt = int(datediff("d", fir1_date, tar1_date)) + 1
			pay_tab(i,5) = work1_cnt
	    end if
	else '초기 년도 퇴직정산
	    tar2_date = cstr(mid(emp_first_date,1,4) + "-" + "12" + "-" + "31")   
		fir2_yy = cstr(mid(emp_first_date,1,4))
		fir2_md = cstr(mid(emp_first_date,6,2) + "-" + mid(emp_first_date,9,2))
		if fir2_md < "10-01" then
		       tcal2_month = mid(cstr(tar2_date),1,4) + mid(cstr(tar2_date),6,2)
		       i = i + 1
               pay_tab(i,1) = cstr(fir2_yy)
			   pay_tab(i,4) = cstr(tcal2_month)
	           tcal2_month = cstr(int(tcal2_month) - 1)
	           pay_tab(i,3) = cstr(tcal2_month)
	           tcal2_month = cstr(int(tcal2_month) - 1)
	           pay_tab(i,2) = cstr(tcal2_month)
		       tar2_date = cstr(mid(pay_tab(i,4),1,4) + "-" + mid(pay_tab(i,4),5,2) + "-" + "31")
		       fir2_date = cstr(mid(pay_tab(i,2),1,4) + "-" + mid(pay_tab(i,2),5,2) + "-" + "01")
		       work2_cnt = int(datediff("d", fir2_date, tar2_date)) + 1
		       pay_tab(i,5) = work2_cnt
			   year2_cnt = int(datediff("d", emp_first_date, tar2_date)) + 1
	           pay_tab(i,6) = year2_cnt
			   start_date = cstr(mid(emp_first_date,1,4) + "-" + "01" + "-" + "01")
	           end_date = cstr(mid(emp_first_date,1,4) + "-" + "12" + "-" + "31")
	           pay_tab(i,7) = int(datediff("d", start_date, end_date)) + 1
		   else
		       tcal2_month = mid(cstr(emp_first_date),1,4) + mid(cstr(emp_first_date),6,2)
			   m_bigo = int(mid(cstr(emp_first_date),6,2))
			   tcal22_month = cstr(int(tcal2_month) + 1)
			   tcal23_month = cstr(int(tcal22_month) + 1)
			   i = i + 1
			   pay_tab(i,1) = cstr(fir2_yy)
               if m_bigo <= 10 then 
			         pay_tab(i,2) = cstr(tcal2_month)
					 pay_tab(i,3) = cstr(tcal22_month)
					 pay_tab(i,4) = cstr(tcal23_month)
					 tar2_date = cstr(mid(pay_tab(i,4),1,4) + "-" + mid(pay_tab(i,4),5,2) + "-" + "31")
			         fir2_date = cstr(mid(pay_tab(i,2),1,4) + "-" + mid(pay_tab(i,2),5,2) + "-" + fcal_day)
				  elseif m_bigo = 11 then
			                pay_tab(i,3) = cstr(tcal2_month)
							pay_tab(i,4) = cstr(tcal22_month)
							tar2_date = cstr(mid(pay_tab(i,4),1,4) + "-" + mid(pay_tab(i,4),5,2) + "-" + "31")
			                fir2_date = cstr(mid(pay_tab(i,3),1,4) + "-" + mid(pay_tab(i,3),5,2) + "-" + fcal_day)
						 elseif m_bigo = 12 then
			                pay_tab(i,4) = cstr(tcal2_month)
							tar2_date = cstr(mid(pay_tab(i,4),1,4) + "-" + mid(pay_tab(i,4),5,2) + "-" + "31")
			                fir2_date = cstr(mid(pay_tab(i,4),1,4) + "-" + mid(pay_tab(i,4),5,2) + "-" + fcal_day)
			   end if
			   work2_cnt = int(datediff("d", fir2_date, tar2_date)) + 1
			   pay_tab(i,5) = work2_cnt
			   year2_cnt = int(datediff("d", emp_first_date, tar2_date)) + 1
	           pay_tab(i,6) = year2_cnt
			   start_date = cstr(mid(emp_first_date,1,4) + "-" + "01" + "-" + "01")
	           end_date = cstr(mid(emp_first_date,1,4) + "-" + "12" + "-" + "31")
	           pay_tab(i,7) = int(datediff("d", start_date, end_date)) + 1
	     end if
		 cf_year = int(mid(cstr(cf_date),1,4)) ' 퇴직연도까지 
		 for i = 2 to year_cnt
				cf_year = cf_year + 1
				pay_tab(i,1) = cstr(cf_year)
				cen_date = cstr(cf_year) + "-" + "01" + "-" + "01"
				start_date = cstr(cf_year) + "-" + "01" + "-" + "01"
	            end_date = cstr(cf_year) + "-" + "12" + "-" + "31"
	            pay_tab(i,7) = int(datediff("d", start_date, end_date)) + 1
				if cf_year <> t_year then 
		               tcal3_month = cstr(cstr(cf_year) + "12")
                       pay_tab(i,4) = cstr(tcal3_month)
	                   tcal3_month = cstr(int(tcal3_month) - 1)
	                   pay_tab(i,3) = cstr(tcal3_month)
	                   tcal3_month = cstr(int(tcal3_month) - 1)
	                   pay_tab(i,2) = cstr(tcal3_month)
	                   tar3_date = cstr(mid(pay_tab(i,4),1,4) + "-" + mid(pay_tab(i,4),5,2) + "-" + "31")
	                   fir3_date = cstr(mid(pay_tab(i,2),1,4) + "-" + mid(pay_tab(i,2),5,2) + "-" + "01")
	                   work3_cnt = int(datediff("d", fir3_date, tar3_date)) + 1
	                   pay_tab(i,5) = work3_cnt
					   year3_cnt = int(datediff("d", cen_date, tar3_date)) + 1
	                   pay_tab(i,6) = year3_cnt
				   else
				       m_bigo = int(mid(cstr(target_date),6,2))
					   tcal3_month = mid(cstr(target_date),1,4) + mid(cstr(target_date),6,2)
					   tcal32_month = cstr(int(tcal3_month) - 1)
			           tcal33_month = cstr(int(tcal32_month) - 1)
                       if m_bigo >= 3 then 
			                 pay_tab(i,4) = cstr(tcal3_month)
				        	 pay_tab(i,3) = cstr(tcal32_month)
				         	 pay_tab(i,2) = cstr(tcal33_month)
							 tar3_date = cstr(mid(pay_tab(i,4),1,4) + "-" + mid(pay_tab(i,4),5,2) + "-" + tcal_day)
			                 fir3_date = cstr(mid(pay_tab(i,2),1,4) + "-" + mid(pay_tab(i,2),5,2) + "-" + "01")
				          elseif m_bigo = 2 then
			                        pay_tab(i,4) = cstr(tcal3_month)
						         	pay_tab(i,3) = cstr(tcal32_month)
									tar3_date = cstr(mid(pay_tab(i,4),1,4) + "-" + mid(pay_tab(i,4),5,2) + "-" + tcal_day)
			                        fir3_date = cstr(mid(pay_tab(i,3),1,4) + "-" + mid(pay_tab(i,3),5,2) + "-" + "01")
						         elseif m_bigo = 1 then
			                        pay_tab(i,2) = cstr(tcal3_month)
									tar3_date = cstr(mid(pay_tab(i,4),1,4) + "-" + mid(pay_tab(i,4),5,2) + "-" + tcal_day)
			                        fir3_date = cstr(mid(pay_tab(i,4),1,4) + "-" + mid(pay_tab(i,4),5,2) + "-" + "01")
			           end if
			           work3_cnt = int(datediff("d", fir3_date, tar3_date)) + 1
			           pay_tab(i,5) = work3_cnt
					   year3_cnt = int(datediff("d", cen_date, tar3_date)) + 1
	                   pay_tab(i,6) = year3_cnt
				end if
		 next
    end if
end if		 

do_i = i

'response.write(fir2_date)
'response.write("/")
'response.write(tar2_date)
'response.write("/")
'response.write(work2_cnt)

for i = 1 to do_i
    for j = 2 to 4
	    p_yymm = pay_tab(i,j)
		if p_yymm <> "" then
		      Sql = "select * from pay_month_give where (pmg_yymm = '"+p_yymm+"' ) and (pmg_id = '1') and (pmg_emp_no = '"+emp_no+"') and (pmg_company = '"+view_condi+"')"
              Rs_give.Open Sql, Dbconn, 1
              Set Rs_give = DbConn.Execute(SQL)
              if not Rs_give.eof then
                     pmg_give_tot = int(Rs_give("pmg_give_total"))	
					 pmg_tax_yes = int(Rs_give("pmg_tax_yes"))	
					 pmg_base_pay = int(Rs_give("pmg_base_pay"))	
					 pmg_meals_pay = int(Rs_give("pmg_meals_pay"))	
					 pmg_overtime_pay = int(Rs_give("pmg_overtime_pay"))	
                 else
                     pmg_give_tot = 0
					 pmg_tax_yes = 0
					 pmg_base_pay = 0
					 pmg_meals_pay = 0
					 pmg_overtime_pay = 0
              end if
			  Rs_give.close()
			  Sql = "select * from pay_month_deduct where (de_yymm = '"+p_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
              Set Rs_dct = DbConn.Execute(SQL)
              if not Rs_dct.eof then
                     de_deduct_tot = int(Rs_dct("de_deduct_total"))	
                 else
                     de_deduct_tot = 0
              end if
              Rs_dct.close()
			  pay_curr_amt = pmg_base_pay + pmg_meals_pay + pmg_overtime_pay
			  pay_pay(i,j) = pay_curr_amt
	     end if
    next
next

'상여금
'for i = 1 to do_i
'    for j = 2 to 4
'	    p_yymm = pay_tab(i,j)
'		if p_yymm <> "" then
'		      Sql = "select * from pay_month_give where (pmg_yymm = '"+p_yymm+"' ) and (pmg_id = '2') and (pmg_emp_no = '"+emp_no+"') and (pmg_company = '"+view_condi+"')"
'              Rs_give.Open Sql, Dbconn, 1
'              Set Rs_give = DbConn.Execute(SQL)
'              if not Rs_give.eof then
'                     pmg_give_tot = int(Rs_give("pmg_give_total"))	
'                 else
'                     pmg_give_tot = 0
'              end if
'			  Rs_give.close()
'			  Sql = "select * from pay_month_deduct where (de_yymm = '"+p_yymm+"' ) and (de_id = '2') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
'              Set Rs_dct = DbConn.Execute(SQL)
'              if not Rs_dct.eof then
'                     de_deduct_tot = int(Rs_dct("de_deduct_total"))	
'                 else
'                     de_deduct_tot = 0
'              end if
'              Rs_dct.close()
'			  pay_curr_amt = pmg_give_tot - de_deduct_tot
'			  bonus_tab(i,j) = pay_curr_amt
'	     end if
'   next
'next

Sql = "SELECT * FROM pay_year_end_tax WHERE end_emp_no = '"+emp_no+"' and end_year = '"+end_year+"' and end_id = '2' and end_pay_type = '1' and end_company = '"+view_condi+"'"
Set Rs_ytax=Dbconn.Execute(sql)
if not Rs_ytax.eof then
       end_add_tax = int(Rs_ytax("end_add_tax"))	
	   end_add_wetax = int(Rs_ytax("end_add_wetax"))
	   end_add_nhis = int(Rs_ytax("end_add_nhis"))
	   end_add_longcare = int(Rs_ytax("end_add_longcare"))	
   else
       end_add_tax = 0
	   end_add_wetax = 0
	   end_add_nhis = 0
	   end_add_longcare = 0
end if
Rs_ytax.close()
add_tax_hap = end_add_tax + end_add_wetax
add_nhis_hap = end_add_nhis + end_add_longcare


title_line = "퇴직급여 처리/입력"

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
				a=confirm('퇴직급여처리를 하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

		    function num_chk(txtObj){
				sum_end_pay = parseInt(document.frm.eos_sum_end_pay.value.replace(/,/g,""));
				incomm_tax = parseInt(document.frm.eos_incomm_tax.value.replace(/,/g,""));
				we_tax = parseInt(document.frm.eos_we_tax.value.replace(/,/g,""));
				loan_jan = parseInt(document.frm.eos_loan_jan.value.replace(/,/g,""));
				loan_interest = parseInt(document.frm.eos_loan_interest.value.replace(/,/g,""));
				nhis_amt = parseInt(document.frm.eos_nhis_amt.value.replace(/,/g,""));

				tax_hap = incomm_tax + we_tax;
				loan_hap = loan_jan + loan_interest;
		        give_amt = sum_end_pay - tax_hap - loan_hap - nhis_amt;
		
				incomm_tax = String(incomm_tax);
				num_len = incomm_tax.length;
				sil_len = num_len;
				incomm_tax = String(incomm_tax);
				if (incomm_tax.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) incomm_tax = incomm_tax.substr(0,num_len -3) + "," + incomm_tax.substr(num_len -3,3);
				if (sil_len > 6) incomm_tax = incomm_tax.substr(0,num_len -6) + "," + incomm_tax.substr(num_len -6,3) + "," + incomm_tax.substr(num_len -2,3);
				document.frm.eos_incomm_tax.value = incomm_tax; 
				
				we_tax = String(we_tax);
				num_len = we_tax.length;
				sil_len = num_len;
				we_tax = String(we_tax);
				if (we_tax.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) we_tax = we_tax.substr(0,num_len -3) + "," + we_tax.substr(num_len -3,3);
				if (sil_len > 6) we_tax = we_tax.substr(0,num_len -6) + "," + we_tax.substr(num_len -6,3) + "," + we_tax.substr(num_len -2,3);
				document.frm.eos_we_tax.value = we_tax; 
				
				tax_hap = String(tax_hap);
				num_len = tax_hap.length;
				sil_len = num_len;
				tax_hap = String(tax_hap);
				if (tax_hap.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) tax_hap = tax_hap.substr(0,num_len -3) + "," + tax_hap.substr(num_len -3,3);
				if (sil_len > 6) tax_hap = tax_hap.substr(0,num_len -6) + "," + tax_hap.substr(num_len -6,3) + "," + tax_hap.substr(num_len -2,3);
				document.frm.eos_tax_hap.value = tax_hap; 
				

				loan_jan = String(loan_jan);
				num_len = loan_jan.length;
				sil_len = num_len;
				loan_jan = String(loan_jan);
				if (loan_jan.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) loan_jan = loan_jan.substr(0,num_len -3) + "," + loan_jan.substr(num_len -3,3);
				if (sil_len > 6) loan_jan = loan_jan.substr(0,num_len -6) + "," + loan_jan.substr(num_len -6,3) + "," + loan_jan.substr(num_len -2,3);
				document.frm.eos_loan_jan.value = loan_jan; 		

				loan_interest = String(loan_interest);
				num_len = loan_interest.length;
				sil_len = num_len;
				loan_interest = String(loan_interest);
				if (loan_interest.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) loan_interest = loan_interest.substr(0,num_len -3) + "," + loan_interest.substr(num_len -3,3);
				if (sil_len > 6) loan_interest = loan_interest.substr(0,num_len -6) + "," + loan_interest.substr(num_len -6,3) + "," + loan_interest.substr(num_len -2,3);
				document.frm.eos_loan_interest.value = loan_interest; 	
		
				loan_hap = String(loan_hap);
				num_len = loan_hap.length;
				sil_len = num_len;
				loan_hap = String(loan_hap);
				if (loan_hap.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) loan_hap = loan_hap.substr(0,num_len -3) + "," + loan_hap.substr(num_len -3,3);
				if (sil_len > 6) loan_hap = loan_hap.substr(0,num_len -6) + "," + loan_hap.substr(num_len -6,3) + "," + loan_hap.substr(num_len -2,3);
				document.frm.eos_loan_hap.value = loan_hap;
				
		
				nhis_amt = String(nhis_amt);
				num_len = nhis_amt.length;
				sil_len = num_len;
				nhis_amt = String(nhis_amt);
				if (nhis_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) nhis_amt = nhis_amt.substr(0,num_len -3) + "," + nhis_amt.substr(num_len -3,3);
				if (sil_len > 6) nhis_amt = nhis_amt.substr(0,num_len -6) + "," + nhis_amt.substr(num_len -6,3) + "," + nhis_amt.substr(num_len -2,3);
				document.frm.eos_nhis_amt.value = nhis_amt;
				
				end_pension = parseInt(document.frm.eos_end_pension.value.replace(/,/g,""));
				curr_give_amt = give_amt - end_pension;
				
				give_amt = String(give_amt);   
				num_len = give_amt.length;
				sil_len = num_len;
				give_amt = String(give_amt);
				if (give_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) give_amt = give_amt.substr(0,num_len -3) + "," + give_amt.substr(num_len -3,3);
				if (sil_len > 6) give_amt = give_amt.substr(0,num_len -6) + "," + give_amt.substr(num_len -6,3) + "," + give_amt.substr(num_len -2,3);
				document.frm.eos_give_amt.value = give_amt;
				
				end_pension = String(end_pension);
				num_len = end_pension.length;
				sil_len = num_len;
				end_pension = String(end_pension);
				if (end_pension.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) end_pension = end_pension.substr(0,num_len -3) + "," + end_pension.substr(num_len -3,3);
				if (sil_len > 6) end_pension = end_pension.substr(0,num_len -6) + "," + end_pension.substr(num_len -6,3) + "," + end_pension.substr(num_len -2,3);
				document.frm.eos_end_pension.value = end_pension;
				
				curr_give_amt = String(curr_give_amt);
				num_len = curr_give_amt.length;
				sil_len = num_len;
				curr_give_amt = String(curr_give_amt);
				if (curr_give_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) curr_give_amt = curr_give_amt.substr(0,num_len -3) + "," + curr_give_amt.substr(num_len -3,3);
				if (sil_len > 6) curr_give_amt = curr_give_amt.substr(0,num_len -6) + "," + curr_give_amt.substr(num_len -6,3) + "," + curr_give_amt.substr(num_len -2,3);
				document.frm.eos_curr_give_amt.value = curr_give_amt;
			
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
				<form action="insa_pay_empout_save.asp" method="post" name="frm">
               	<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>퇴직대상</dt>
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
                            <col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
							<col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
							<col width="9%" > 
                            <col width="9%" >
                            <col width="9%" >
                            <col width="7%" >
                            <col width="4%" >
						</colgroup>
                        <thead>
				            <tr>
				               <th class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">년도<br>근무일수</th>
				               <th scope="col" style=" border-bottom:1px solid #e3e3e3;">구분</th>
                               <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">최근3개월급여</th>
				               <th scope="col" style=" border-bottom:1px solid #e3e3e3;">총일수</th>
                               <th scope="col" style=" border-bottom:1px solid #e3e3e3;">합계</th>
                               <th scope="col" style=" border-bottom:1px solid #e3e3e3;">총액<br>(급여+상여)</th>
                               <th scope="col" style=" border-bottom:1px solid #e3e3e3;">평균임금</th>
                               <th scope="col" style=" border-bottom:1px solid #e3e3e3;">퇴직금</th>
                               <th scope="col" style=" border-bottom:1px solid #e3e3e3;">비고</th>
                               <th scope="col" style=" border-bottom:1px solid #e3e3e3;">변경</th>
			               </tr>
						</thead>
						<tbody>
                        <%
						sum_end_pay = 0
                        for i = 1 to 50 
                        	if	pay_tab(i,3) <> "" then
								pay_sum = pay_pay(i,2)+pay_pay(i,3)+pay_pay(i,4)+bonus_tab(i,2)+bonus_tab(i,3)+bonus_tab(i,4)
							    eot_average_pay = int(pay_sum / pay_tab(i,5))
							    eot_end_pay = eot_average_pay * 30 * pay_tab(i,6) / pay_tab(i,7)
								sum_end_pay = sum_end_pay + eot_end_pay
						%>	
							<tr>
								<td rowspan="3" class="left"><%=mid(pay_tab(i,1),1,4)%>년&nbsp;(<%=pay_tab(i,6)%>)</td>
								<td class="left" ><%=eot_space%>&nbsp;</td>
                                <td><%=mid(pay_tab(i,2),1,4)%>년&nbsp;<%=mid(pay_tab(i,2),5,2)%>월</td>
                                <td><%=mid(pay_tab(i,3),1,4)%>년&nbsp;<%=mid(pay_tab(i,3),5,2)%>월</td>
                                <td><%=mid(pay_tab(i,4),1,4)%>년&nbsp;<%=mid(pay_tab(i,4),5,2)%>월</td>
                                <td rowspan="3" class="right"><%=pay_tab(i,5)%>&nbsp;-&nbsp;<%=pay_tab(i,7)%>&nbsp;</td>
                                <td class="left" ><%=eot_space%>&nbsp;</td>
                                <td rowspan="3" class="right"><%=formatnumber(clng(pay_pay(i,2)+pay_pay(i,3)+pay_pay(i,4)+bonus_tab(i,2)+bonus_tab(i,3)+bonus_tab(i,4)),0)%>&nbsp;</td>
                                <td rowspan="3" class="right"><%=formatnumber(eot_average_pay,0)%>&nbsp;</td>
                                <td rowspan="3" class="right"><%=formatnumber(eot_end_pay,0)%>&nbsp;</td>
                                <td rowspan="3" class="left"><%=eot_comment%>&nbsp;</td>
                                <td rowspan="3" class="left">수정</td>
							</tr>
                            <tr>
								<td style=" border-left:1px solid #e3e3e3;" >급여</td>
                                <td class="right" ><%=formatnumber(pay_pay(i,2),0)%>&nbsp;</td>
                                <td class="right" ><%=formatnumber(pay_pay(i,3),0)%>&nbsp;</td>
                                <td class="right" ><%=formatnumber(pay_pay(i,4),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(clng(pay_pay(i,2)+pay_pay(i,3)+pay_pay(i,4)),0)%>&nbsp;</td>
							</tr>
   							<tr>
								<td style=" border-left:1px solid #e3e3e3;" >상여</td>
                                <td class="right" ><%=formatnumber(bonus_tab(i,2),0)%>&nbsp;</td>
                                <td class="right" ><%=formatnumber(bonus_tab(i,3),0)%>&nbsp;</td>
                                <td class="right" ><%=formatnumber(bonus_tab(i,4),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(clng(bonus_tab(i,2)+bonus_tab(i,3)+bonus_tab(i,4)),0)%>&nbsp;</td>
							</tr>
                      <%
							end if
						next
                      %>
                            <tr>
				               <th rowspan="2" class="first" scope="col" style=" border-top:2px solid #515254;">퇴직금합계</th>
                               <th colspan="7" scope="col" style=" border-top:2px solid #515254;">기타 미정산 내역</th>
				               <th rowspan="2" scope="col" style=" border-top:2px solid #515254;">지급총액</th>
                               <th rowspan="2" scope="col" style=" border-top:2px solid #515254;">기퇴직연금<br>입금액</th>
                               <th rowspan="2" colspan="2" scope="col" style=" border-top:2px solid #515254;">총실지급액</th>
			               </tr>
                           <tr>
				               <th scope="col" style=" border-left:1px solid #e3e3e3;">소득세</th>
                               <th scope="col">주민세</th>
                               <th scope="col">계</th>
                               <th scope="col">대출잔액</th>
                               <th scope="col">대출이자</th>
                               <th scope="col">계</th>
                               <th scope="col">건강보험</th>
			               </tr>
                           <tr>
								<td rowspan="2" class="right"><%=formatnumber(sum_end_pay,0)%>&nbsp;
                                <input name="eos_sum_end_pay" type="hidden" id="eos_sum_end_pay" style="width:90px;text-align:right" value="<%=formatnumber(sum_end_pay,0)%>">
                                </td>
								<td class="left" >
                                <input name="eos_incomm_tax" type="text" id="eos_incomm_tax" style="width:90px;text-align:right" value="<%=formatnumber(end_add_tax,0)%>" onKeyUp="num_chk(this);">&nbsp;
                                </td>
                                <td class="left" >
                                <input name="eos_we_tax" type="text" id="eos_we_tax" style="width:90px;text-align:right" value="<%=formatnumber(end_add_wetax,0)%>" onKeyUp="num_chk(this);">&nbsp;
                                </td>
                                <td class="left" >
                                <input name="eos_tax_hap" type="text" id="eos_tax_hap" style="width:90px;text-align:right" value="<%=formatnumber(add_tax_hap,0)%>" readonly="true">&nbsp;
                                </td>
                                <td class="left" >
                                <input name="eos_loan_jan" type="text" id="eos_loan_jan" style="width:90px;text-align:right" value="<%=formatnumber(eos_loan_jan,0)%>" onKeyUp="num_chk(this);">&nbsp;
                                </td>
                                <td class="left" >
                                <input name="eos_loan_interest" type="text" id="eos_loan_interest" style="width:90px;text-align:right" value="<%=formatnumber(eos_loan_interest,0)%>" onKeyUp="num_chk(this);">&nbsp;
                                </td>
                                <td class="left" >
                                <input name="eos_loan_hap" type="text" id="eos_loan_hap" style="width:90px;text-align:right" value="<%=formatnumber(eos_loan_hap,0)%>" readonly="true">&nbsp;
                                </td>
                                <td class="left" >
                                <input name="eos_nhis_amt" type="text" id="eos_nhis_amt" style="width:90px;text-align:right" value="<%=formatnumber(add_nhis_hap,0)%>" onKeyUp="num_chk(this);">&nbsp;
                                </td>
                                <td rowspan="2" class="left" >
                                <input name="eos_give_amt" type="text" id="eos_give_amt" style="width:90px;text-align:right" value="<%=formatnumber(eos_give_amt,0)%>" readonly="true">&nbsp;
                                </td>
                                <td rowspan="2" class="left" >
                                <input name="eos_end_pension" type="text" id="eos_end_pension" style="width:90px;text-align:right" value="<%=formatnumber(eos_end_pension,0)%>" onKeyUp="num_chk(this);">&nbsp;
                                </td>
                                <td rowspan="2" colspan="2" class="left" >
                                <input name="eos_curr_give_amt" type="text" id="eos_curr_give_amt" style="width:90px;text-align:right" value="<%=formatnumber(eos_curr_give_amt,0)%>" readonly="true">&nbsp;
                                </td>
							</tr>
                      </tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="퇴직급여처리" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="eot_company" value="<%=emp_company%>" ID="Hidden1">
                <input type="hidden" name="eot_bonbu" value="<%=emp_bonbu%>" ID="Hidden1">
                <input type="hidden" name="eot_saupbu" value="<%=emp_saupbu%>" ID="Hidden1">
                <input type="hidden" name="eot_team" value="<%=emp_team%>" ID="Hidden1">
                <input type="hidden" name="eot_org_name" value="<%=emp_org_name%>" ID="Hidden1">
                <input type="hidden" name="eot_org_code" value="<%=emp_org_code%>" ID="Hidden1">
			</form> 
		</div>				
	</body>
</html>

