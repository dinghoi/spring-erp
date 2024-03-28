<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next
dim pay_tab(50,10)
dim pay_pay(50,10)
dim bonus_tab(50,10)

u_type = request.form("u_type")
emp_no = request.form("emp_no")

for i = 1 to 50
    for j = 1 to 10
	    pay_tab(i,j) = ""
		pay_pay(i,j) = 0
		bonus_tab(i,j) = 0
    next
next

set dbconn = server.CreateObject("adodb.connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_eout = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect
	
Sql = "SELECT * FROM emp_master where emp_no = '"+emp_no+"'"
Set rs_emp = DbConn.Execute(SQL)
if not rs_emp.eof then
    	emp_name = rs_emp("emp_name")
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
		      Sql = "select * from pay_month_give where (pmg_yymm = '"+p_yymm+"' ) and (pmg_id = '1') and (pmg_emp_no = '"+emp_no+"')"
              Rs_give.Open Sql, Dbconn, 1
              Set Rs_give = DbConn.Execute(SQL)
              if not Rs_give.eof then
                     pmg_give_tot = int(Rs_give("pmg_give_total"))	
                 else
                     pmg_give_tot = 0
              end if
			  Rs_give.close()
			  Sql = "select * from pay_month_deduct where (de_yymm = '"+p_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"')"
              Set Rs_dct = DbConn.Execute(SQL)
              if not Rs_dct.eof then
                     de_deduct_tot = int(Rs_dct("de_deduct_total"))	
                 else
                     de_deduct_tot = 0
              end if
              Rs_dct.close()
			  pay_curr_amt = pmg_give_tot - de_deduct_tot
			  pay_pay(i,j) = pay_curr_amt
	     end if
    next
next

for i = 1 to do_i
    for j = 2 to 4
	    p_yymm = pay_tab(i,j)
		if p_yymm <> "" then
		      Sql = "select * from pay_month_give where (pmg_yymm = '"+p_yymm+"' ) and (pmg_id = '2') and (pmg_emp_no = '"+emp_no+"')"
              Rs_give.Open Sql, Dbconn, 1
              Set Rs_give = DbConn.Execute(SQL)
              if not Rs_give.eof then
                     pmg_give_tot = int(Rs_give("pmg_give_total"))	
                 else
                     pmg_give_tot = 0
              end if
			  Rs_give.close()
			  Sql = "select * from pay_month_deduct where (de_yymm = '"+p_yymm+"' ) and (de_id = '2') and (de_emp_no = '"+emp_no+"')"
              Set Rs_dct = DbConn.Execute(SQL)
              if not Rs_dct.eof then
                     de_deduct_tot = int(Rs_dct("de_deduct_total"))	
                 else
                     de_deduct_tot = 0
              end if
              Rs_dct.close()
			  pay_curr_amt = pmg_give_tot - de_deduct_tot
			  bonus_tab(i,j) = pay_curr_amt
	     end if
    next
next
	
	eos_sum_end_pay = int(request.form("eos_sum_end_pay"))
	eos_incomm_tax = int(request.form("eos_incomm_tax"))
	eos_we_tax = int(request.form("eos_we_tax"))
	eos_tax_hap = int(request.form("eos_tax_hap"))
	eos_loan_jan = int(request.form("eos_loan_jan"))
	eos_loan_interest = int(request.form("eos_loan_interest"))
	eos_loan_hap = int(request.form("eos_loan_hap"))
	eos_nhis_amt = int(request.form("eos_nhis_amt"))
	eos_give_amt = int(request.form("eos_give_amt"))
	eos_end_pension = int(request.form("eos_end_pension"))
	eos_curr_give_amt = int(request.form("eos_curr_give_amt"))

dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")
sms_msg = emp_no + "-" + emp_name + "- 퇴직급여 처리가 "

Sql = "SELECT * FROM pay_empout_sum WHERE eos_empno = '"+emp_no+"'"
Set Rs_eout=Dbconn.Execute(sql)
if Rs_eout.eof then

   for i = 1 to 50 
   	if	pay_tab(i,3) <> "" then
	    pay_hap = pay_pay(i,2)+pay_pay(i,3)+pay_pay(i,4)
		bonus_hap = bonus_tab(i,2)+bonus_tab(i,3)+bonus_tab(i,4)
		pay_sum = pay_pay(i,2)+pay_pay(i,3)+pay_pay(i,4)+bonus_tab(i,2)+bonus_tab(i,3)+bonus_tab(i,4)
	    eot_average_pay = int(pay_sum / pay_tab(i,5))
	    eot_end_pay = eot_average_pay * 30 * pay_tab(i,6) / pay_tab(i,7)

		sql="insert into pay_empout (eot_empno,eot_end_date,eot_yyyy,eot_emp_name,eot_first_date,eot_in_date,eot_company,eot_bonbu,eot_saupbu,eot_team,eot_org_name,eot_org_code,eot_grade,eot_position,eot_mon1_day,eot_mon2_day,eot_mon3_day,eot_mon1,eot_mon2,eot_mon3,eot_bonus,eot_year_day,eot_work_cnt,eot_work_day,eot_mon_hap,eot_bonus_hap,eot_pay_sum,eot_average_pay,eot_end_pay,eot_comment,eot_reg_date,eot_reg_user) values ('"&emp_no&"','"&emp_end_date&"','"&mid(pay_tab(i,1),1,4)&"','"&emp_name&"','"&emp_first_date&"','"&emp_in_date&"','"&emp_company&"','"&emp_bonbu&"','"&emp_saupbu&"','"&emp_team&"','"&emp_org_name&"','"&emp_org_code&"','"&emp_grade&"','"&emp_position&"','"&mid(pay_tab(i,2),1,6)&"','"&mid(pay_tab(i,3),1,6)&"','"&mid(pay_tab(i,4),1,6)&"','"&pay_pay(i,2)&"','"&pay_pay(i,3)&"','"&pay_pay(i,4)&"','"&bonus_hap&"','"&pay_tab(i,7)&"','"&pay_tab(i,5)&"','"&pay_tab(i,6)&"','"&pay_hap&"','"&bonus_hap&"','"&pay_sum&"','"&eot_average_pay&"','"&eot_end_pay&"','',now(),'"&emp_user&"')"
		
		dbconn.execute(sql)
		
    end if
  next	

		sql="insert into pay_empout_sum (eos_empno,eos_end_date,eos_emp_name,eos_first_date,eos_in_date,eos_company,eos_bonbu,eos_saupbu,eos_team,eos_org_name,eos_org_code,eos_grade,eos_position,eos_sum_end_pay,eos_incomm_tax,eos_we_tax,eos_tax_hap,eos_loan_jan,eos_loan_interest,eos_loan_hap,eos_nhis_amt,eos_give_amt,eos_end_pension,eos_curr_give_amt,eos_reg_date,eos_reg_user) values ('"&emp_no&"','"&emp_end_date&"','"&emp_name&"','"&emp_first_date&"','"&emp_in_date&"','"&emp_company&"','"&emp_bonbu&"','"&emp_saupbu&"','"&emp_team&"','"&emp_org_name&"','"&emp_org_code&"','"&emp_grade&"','"&emp_position&"','"&eos_sum_end_pay&"','"&eos_incomm_tax&"','"&eos_we_tax&"','"&eos_tax_hap&"','"&eos_loan_jan&"','"&eos_loan_interest&"','"&eos_loan_hap&"','"&eos_nhis_amt&"','"&eos_give_amt&"','"&eos_end_pension&"','"&eos_curr_give_amt&"',now(),'"&emp_user&"')"
		dbconn.execute(sql)
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "저장중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + "저장되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"location.replace('insa_pay_empout_mg.asp');"
	'response.write"self.close() ;"
	response.write"</script>"
	Response.End

  else
	response.write"<script language=javascript>"
	response.write"alert('이미 퇴직급여처리를 하였습니다..');"		
	response.write"location.replace('insa_pay_empout_mg.asp');"
	response.write"</script>"
	Response.End
end if	

	dbconn.Close()
	Set dbconn = Nothing

%>
