<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

emp_user = request.cookies("nkpmg_user")("coo_user_name")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

view_condi=Request("view_condi")
from_date=request("from_date")
to_date=request("to_date")
pmg_yymm=request("pmg_yymm")

'response.write(pmg_yymm)
'response.write(view_condi)
'response.write(pmg_yymm_to)
'response.write(to_date)
'response.End

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_this = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

dbconn.BeginTrans

pmg_id = "1"

Sql = "select * from pay_overtime_cost where emp_company = '"+view_condi+"' and work_date >= '"+from_date+"' and work_date <= '"+to_date+"' ORDER BY emp_company,team,org_name,mg_ce_id,work_date ASC"
Rs.Open Sql, Dbconn, 1

sum_overtime_cnt = 0	 
sum_overtime_cost = 0
							 
tot_overtime_cnt = 0	 
tot_overtime_cost = 0

i = 0
j = 0

if rs.eof or rs.bof then
 	    bi_team = ""
		bi_ce = ""
   else						  
		if isnull(rs("team")) or rs("team") = "" then	
				bi_team = ""
  		   else
				bi_team = rs("team")
		end if
		if isnull(rs("mg_ce_id")) or rs("mg_ce_id") = "" then	
				bi_mg_ce_id = ""
		   else
				bi_mg_ce_id = rs("mg_ce_id")
		end if
end if

do until rs.eof
   
   view_condi = rs("emp_company")
   
   if isnull(rs("team")) or rs("team") = "" then
		  emp_team = ""
      else
  	      emp_team = rs("team")
   end if
   if isnull(rs("mg_ce_id")) or rs("mg_ce_id") = "" then
	      mg_ce_id = ""
	  else
	      mg_ce_id = rs("mg_ce_id")
   end if
   							
   if bi_mg_ce_id <> mg_ce_id then
	     emp_no = bi_mg_ce_id
	     Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
         Set rs_emp = DbConn.Execute(SQL)
	     if not Rs_emp.eof then
		      ce_name = rs_emp("emp_name")
	     end if
	     rs_emp.close()

         Sql = "SELECT * FROM pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '"+pmg_id+"') and (pmg_emp_no = '"+bi_mg_ce_id+"') and (pmg_company = '"+view_condi+"')"
         Set Rs_give = DbConn.Execute(SQL)
	     if not Rs_give.eof then
		      pmg_base_pay = Rs_give("pmg_base_pay")
	          pmg_meals_pay = Rs_give("pmg_meals_pay")
	          pmg_postage_pay = Rs_give("pmg_postage_pay")
	          pmg_re_pay = Rs_give("pmg_re_pay")
	          pmg_overtime_pay = Rs_give("pmg_overtime_pay")
	          pmg_car_pay = Rs_give("pmg_car_pay")
	          pmg_position_pay = Rs_give("pmg_position_pay")
	          pmg_custom_pay = Rs_give("pmg_custom_pay")
	          pmg_job_pay = Rs_give("pmg_job_pay")
			  
	          pmg_job_support = sum_overtime_cost
			  
	          pmg_jisa_pay = Rs_give("pmg_jisa_pay")
	          pmg_long_pay = Rs_give("pmg_long_pay")
	          pmg_disabled_pay = Rs_give("pmg_disabled_pay")
	          pmg_family_pay = Rs_give("pmg_family_pay")
	          pmg_school_pay = Rs_give("pmg_school_pay")
	          pmg_qual_pay = Rs_give("pmg_qual_pay")
	          pmg_other_pay1 = Rs_give("pmg_other_pay1")
	          pmg_other_pay2 = Rs_give("pmg_other_pay2")
	          pmg_other_pay3 = Rs_give("pmg_other_pay3")
	          pmg_tax_yes = Rs_give("pmg_tax_yes")
	          pmg_tax_no = Rs_give("pmg_tax_no")
	          pmg_tax_reduced = Rs_give("pmg_tax_reduced")
			  
			  meals_taxno_pay = pmg_meals_pay
	          car_taxno_pay = pmg_car_pay
	          meals_tax_pay = 0
	          car_tax_pay = 0
	          if (meals_pay > 100000) then
	                meals_tax_pay = Int(meals_pay - 100000)
	          end if
	          if (meals_pay > 100000) then 
	                meals_taxno_pay =  100000
	          end if
	          if (car_pay > 200000) then
	                car_tax_pay = Int(car_pay - 200000)
	          end if
	          if (car_pay > 200000) then
	                car_taxno_pay =  200000
	          end if
				
	          pmg_tax_yes = pmg_base_pay + pmg_postage_pay + pmg_re_pay + pmg_overtime_pay + pmg_position_pay + pmg_custom_pay + pmg_job_pay + pmg_job_support + pmg_jisa_pay + pmg_long_pay + pmg_disabled_pay + meals_tax_pay + car_tax_pay
				
	          pmg_tax_no = meals_taxno_pay + car_taxno_pay
	
	          pmg_give_tot = pmg_tax_yes + pmg_tax_no
			  
			  
			  sql = "update pay_month_give set pmg_job_support='"&sum_overtime_cost&"',pmg_tax_yes='"&pmg_tax_yes&"',pmg_tax_no='"&pmg_tax_no&"',pmg_give_total='"&pmg_give_tot&"' where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '"+pmg_id+"') and (pmg_emp_no = '"+bi_mg_ce_id+"') and (pmg_company = '"+view_condi+"')"
		
		      dbconn.execute(sql) 
			  
			  i = i + 1
			else
			  j = j + 1
	     end if
	     Rs_give.close()
		 
         sum_overtime_cnt = 0	 
		 sum_overtime_cost = 0
		 bi_mg_ce_id = mg_ce_id
   end if

   emp_no = rs("mg_ce_id")
   Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
   Set rs_emp = DbConn.Execute(SQL)
   if not Rs_emp.eof then
          emp_company = rs_emp("emp_company")
   	      emp_name = rs_emp("emp_name")
		  emp_end_date = rs_emp("emp_end_date")
   end if
   rs_emp.close()
                          
   if isNull(emp_end_date) or emp_end_date = "1900-01-01" then
          emp_end = ""
  	  else 
	      emp_end = "퇴직"
   end if
							  
   sum_overtime_cnt = sum_overtime_cnt + 1	 
   sum_overtime_cost = sum_overtime_cost + int(rs("overtime_amt"))
   
   rs.movenext()
loop
rs.close()

    Sql = "SELECT * FROM pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '"+pmg_id+"') and (pmg_emp_no = '"+bi_mg_ce_id+"') and (pmg_company = '"+view_condi+"')"
    Set Rs_give = DbConn.Execute(SQL)
	if not Rs_give.eof then
		      pmg_base_pay = Rs_give("pmg_base_pay")
	          pmg_meals_pay = Rs_give("pmg_meals_pay")
	          pmg_postage_pay = Rs_give("pmg_postage_pay")
	          pmg_re_pay = Rs_give("pmg_re_pay")
	          pmg_overtime_pay = Rs_give("pmg_overtime_pay")
	          pmg_car_pay = Rs_give("pmg_car_pay")
	          pmg_position_pay = Rs_give("pmg_position_pay")
	          pmg_custom_pay = Rs_give("pmg_custom_pay")
	          pmg_job_pay = Rs_give("pmg_job_pay")
			  
	          pmg_job_support = sum_overtime_cost
			  
	          pmg_jisa_pay = Rs_give("pmg_jisa_pay")
	          pmg_long_pay = Rs_give("pmg_long_pay")
	          pmg_disabled_pay = Rs_give("pmg_disabled_pay")
	          pmg_family_pay = Rs_give("pmg_family_pay")
	          pmg_school_pay = Rs_give("pmg_school_pay")
	          pmg_qual_pay = Rs_give("pmg_qual_pay")
	          pmg_other_pay1 = Rs_give("pmg_other_pay1")
	          pmg_other_pay2 = Rs_give("pmg_other_pay2")
	          pmg_other_pay3 = Rs_give("pmg_other_pay3")
	          pmg_tax_yes = Rs_give("pmg_tax_yes")
	          pmg_tax_no = Rs_give("pmg_tax_no")
	          pmg_tax_reduced = Rs_give("pmg_tax_reduced")
			  
			  meals_taxno_pay = pmg_meals_pay
	          car_taxno_pay = pmg_car_pay
	          meals_tax_pay = 0
	          car_tax_pay = 0
	          if (meals_pay > 100000) then
	                meals_tax_pay = Int(meals_pay - 100000)
	          end if
	          if (meals_pay > 100000) then 
	                meals_taxno_pay =  100000
	          end if
	          if (car_pay > 200000) then
	                car_tax_pay = Int(car_pay - 200000)
	          end if
	          if (car_pay > 200000) then
	                car_taxno_pay =  200000
	          end if
				
	          pmg_tax_yes = pmg_base_pay + pmg_postage_pay + pmg_re_pay + pmg_overtime_pay + pmg_position_pay + pmg_custom_pay + pmg_job_pay + pmg_job_support + pmg_jisa_pay + pmg_long_pay + pmg_disabled_pay + meals_tax_pay + car_tax_pay
				
	          pmg_tax_no = meals_taxno_pay + car_taxno_pay
	
	          pmg_give_tot = pmg_tax_yes + pmg_tax_no
			  
			  
			  sql = "update pay_month_give set pmg_job_support='"&sum_overtime_cost&"',pmg_tax_yes='"&pmg_tax_yes&"',pmg_tax_no='"&pmg_tax_no&"',pmg_give_total='"&pmg_give_tot&"' where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '"+pmg_id+"') and (pmg_emp_no = '"+bi_mg_ce_id+"') and (pmg_company = '"+view_condi+"')"	      
		  
	          dbconn.execute(sql) 
			  
		      i = i + 1
		    else
		      j = j + 1
	    end if
	    Rs_give.close()
		 
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "저장중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		if i = 0 and j > 0 then
		    end_msg = "야.특근 수당 미등록(급여이월부터 하시기 바랍니다)..->. "&j&" 건."
		end if
		if i > 0 and j > 0 then
		    end_msg = "야.특근 수당 등록..->. "&i&" 건....미등록(급여자료가 없습니다)..->. "&j&" 건."
		end if
		if i > 0 and j = 0 then
		    end_msg = "야.특근 수당 등록..->. "&i&" 건"
		end if
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	'response.write"parent.opener.location.reload();"
	response.write"location.replace('insa_pay_overtime_report2.asp');"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
