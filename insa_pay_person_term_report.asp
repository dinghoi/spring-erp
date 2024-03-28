<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

' 기간 조건...sum

Dim Rs
Dim Repeat_Rows
dim month_tab(100,2)

dim com_tab(6)
dim pay_count(6)
dim sum_base_pay(6)
dim sum_meals_pay(6)
dim sum_postage_pay(6)
dim sum_re_pay(6)
dim sum_overtime_pay(6)
dim sum_car_pay(6)
dim sum_position_pay(6)
dim sum_custom_pay(6)
dim sum_job_pay(6)
dim sum_job_support(6)
dim sum_jisa_pay(6)
dim sum_long_pay(6)
dim sum_disabled_pay(65)
dim sum_give_tot(6)

dim sum_nps_amt(6)
dim sum_nhis_amt(6)
dim sum_epi_amt(6)
dim sum_longcare_amt(6)
dim sum_income_tax(6)
dim sum_wetax(6)
dim sum_year_incom_tax(6)
dim sum_year_wetax(6)
dim sum_year_incom_tax2(6)
dim sum_year_wetax2(6)
dim sum_other_amt1(6)
dim sum_sawo_amt(6)
dim sum_johab_amt(6)
dim sum_hyubjo_amt(6)
dim sum_school_amt(6)
dim sum_nhis_bla_amt(6)
dim sum_long_bla_amt(6)
dim sum_deduct_tot(6)

be_pg = "insa_pay_person_term_report.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	f_yymm=Request.form("from_yymm")
	t_yymm=Request.form("to_yymm")
  else
	view_condi = request("view_condi")
	f_yymm=request("from_yymm")
	t_yymm=Request("to_yymm")
end if

if view_condi = "" then
	view_condi = "케이원정보통신"
	curr_dd = cstr(datepart("d",now))
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	from_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	to_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	
	f_yymm = from_yymm
	t_yymm = to_yymm
	
	for i = 1 to 6
        com_tab(i) = ""
        pay_count(i) = 0
        sum_base_pay(i) = 0
        sum_meals_pay(i) = 0
        sum_postage_pay(i) = 0
        sum_re_pay(i) = 0
        sum_overtime_pay(i) = 0
        sum_car_pay(i) = 0
        sum_position_pay(i) = 0
        sum_custom_pay(i) = 0
        sum_job_pay(i) = 0
        sum_job_support(i) = 0
        sum_jisa_pay(i) = 0
        sum_long_pay(i) = 0
        sum_disabled_pay(i) = 0
        sum_give_tot(i) = 0
        sum_nps_amt(i) = 0
        sum_nhis_amt(i) = 0
        sum_epi_amt(i) = 0
        sum_longcare_amt(i) = 0
        sum_income_tax(i) = 0
        sum_wetax(i) = 0
        sum_year_incom_tax(i) = 0
        sum_year_wetax(i) = 0
		sum_year_incom_tax2(i) = 0
        sum_year_wetax2(i) = 0
        sum_other_amt1(i) = 0
        sum_sawo_amt(i) = 0
		sum_johab_amt(i) = 0
        sum_hyubjo_amt(i) = 0
        sum_school_amt(i) = 0
        sum_nhis_bla_amt(i) = 0
        sum_long_bla_amt(i) = 0
        sum_deduct_tot(i) = 0
    next
	
	sum_curr_pay = 0	
	
end if

' 년월 테이블생성
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
month_tab(100,1) = cal_month
view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
month_tab(100,2) = view_month
for i = 1 to 99
	cal_month = cstr(int(cal_month) - 1)
	if mid(cal_month,5) = "00" then
		cal_year = cstr(int(mid(cal_month,1,4)) - 1)
		cal_month = cal_year + "12"
	end if	 
	view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
	j = 100 - i
	month_tab(j,1) = cal_month
	month_tab(j,2) = view_month
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_sum = Server.CreateObject("ADODB.Recordset")
Set Rs_cost = Server.CreateObject("ADODB.Recordset")
Set Rs_old = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY cost_center,cost_group,pmg_saupbu,pmg_org_name,pmg_emp_no ASC"
'order_Sql = " ORDER BY pmg_org_name,pmg_emp_no ASC"
if view_condi = "전체" then
      com_tab(1) = "케이원정보통신"
	  com_tab(2) = "휴디스"
	  com_tab(3) = "케이네트웍스"
	  com_tab(4) = "에스유에이치"
	  com_tab(5) = "코리아디엔씨"
	  com_tab(6) = "합계"
	  where_sql = " WHERE (pmg_yymm >= '"+f_yymm+"' and pmg_yymm <= '"+t_yymm+"') and (pmg_id = '1')" 
   else  
      com_tab(1) = view_condi
	  com_tab(6) = "합계"
	  where_sql = " WHERE (pmg_yymm >= '"+f_yymm+"' and pmg_yymm <= '"+t_yymm+"') and (pmg_id = '1') and (pmg_company = '"+view_condi+"')"
end if   

sql = "select * from pay_month_give " + where_sql + order_sql
Rs.Open Sql, Dbconn, 1
do until rs.eof
    emp_no = rs("pmg_emp_no")
	pmg_company = rs("pmg_company")
	pmg_yymm = rs("pmg_yymm")
				  
    for i = 1 to 6
        if com_tab(i) = rs("pmg_company") then
	             pay_count(i) = pay_count(i) + 1
				 pay_count(6) = pay_count(6) + 1
		         sum_base_pay(i) = sum_base_pay(i) + int(rs("pmg_base_pay"))
                 sum_meals_pay(i) = sum_meals_pay(i) + int(rs("pmg_meals_pay"))
                 sum_postage_pay(i) = sum_postage_pay(i) + int(rs("pmg_postage_pay"))
                 sum_re_pay(i) = sum_re_pay(i) + int(rs("pmg_re_pay"))
                 sum_overtime_pay(i) = sum_overtime_pay(i) + int(rs("pmg_overtime_pay"))
                 sum_car_pay(i) = sum_car_pay(i) + int(rs("pmg_car_pay"))
                 sum_position_pay(i) = sum_position_pay(i) + int(rs("pmg_position_pay"))
                 sum_custom_pay(i) = sum_custom_pay(i) + int(rs("pmg_custom_pay"))
                 sum_job_pay(i) = sum_job_pay(i) + int(rs("pmg_job_pay"))
                 sum_job_support(i) = sum_job_support(i) + int(rs("pmg_job_support"))
                 sum_jisa_pay(i) = sum_jisa_pay(i) + int(rs("pmg_jisa_pay"))
                 sum_long_pay(i) = sum_long_pay(i) + int(rs("pmg_long_pay"))
                 sum_disabled_pay(i) = sum_disabled_pay(i) + int(rs("pmg_disabled_pay"))
                 sum_give_tot(i) = sum_give_tot(i) + int(rs("pmg_give_total"))
				 
				 sum_base_pay(6) = sum_base_pay(6) + int(rs("pmg_base_pay"))
                 sum_meals_pay(6) = sum_meals_pay(6) + int(rs("pmg_meals_pay"))
                 sum_postage_pay(6) = sum_postage_pay(6) + int(rs("pmg_postage_pay"))
                 sum_re_pay(6) = sum_re_pay(6) + int(rs("pmg_re_pay"))
                 sum_overtime_pay(6) = sum_overtime_pay(6) + int(rs("pmg_overtime_pay"))
                 sum_car_pay(6) = sum_car_pay(6) + int(rs("pmg_car_pay"))
                 sum_position_pay(6) = sum_position_pay(6) + int(rs("pmg_position_pay"))
                 sum_custom_pay(6) = sum_custom_pay(6) + int(rs("pmg_custom_pay"))
                 sum_job_pay(6) = sum_job_pay(6) + int(rs("pmg_job_pay"))
                 sum_job_support(6) = sum_job_support(6) + int(rs("pmg_job_support"))
                 sum_jisa_pay(6) = sum_jisa_pay(6) + int(rs("pmg_jisa_pay"))
                 sum_long_pay(6) = sum_long_pay(6) + int(rs("pmg_long_pay"))
                 sum_disabled_pay(6) = sum_disabled_pay(6) + int(rs("pmg_disabled_pay"))
                 sum_give_tot(6) = sum_give_tot(6) + int(rs("pmg_give_total"))
	    end if		 
	next		

    Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+pmg_company+"')"
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
			de_year_incom_tax2 = int(Rs_dct("de_year_incom_tax2"))
            de_year_wetax2 = int(Rs_dct("de_year_wetax2"))
            de_other_amt1 = int(Rs_dct("de_other_amt1"))
            de_sawo_amt = int(Rs_dct("de_sawo_amt"))
            de_johab_amt = int(Rs_dct("de_johab_amt"))
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
			de_year_incom_tax2 = 0
            de_year_wetax2 = 0
            de_other_amt1 = 0
            de_sawo_amt = 0
			de_johab_amt = 0
            de_hyubjo_amt = 0
            de_school_amt = 0
            de_nhis_bla_amt = 0
            de_long_bla_amt = 0
		    de_deduct_tot = 0
     end if
     Rs_dct.close()
     for i = 1 to 6
        if com_tab(i) = rs("pmg_company") then
		         sum_nps_amt(i) = sum_nps_amt(i) + de_nps_amt
                 sum_nhis_amt(i) = sum_nhis_amt(i) + de_nhis_amt
                 sum_epi_amt(i) = sum_epi_amt(i) + de_epi_amt
	             sum_longcare_amt(i) = sum_longcare_amt(i) + de_longcare_amt
                 sum_income_tax(i) = sum_income_tax(i) + de_income_tax
                 sum_wetax(i) = sum_wetax(i) + de_wetax
	             sum_year_incom_tax(i) = sum_year_incom_tax(i) + de_year_incom_tax
                 sum_year_wetax(i) = sum_year_wetax(i) + de_year_wetax
				 sum_year_incom_tax2(i) = sum_year_incom_tax2(i) + de_year_incom_tax2
                 sum_year_wetax2(i) = sum_year_wetax2(i) + de_year_wetax2
                 sum_other_amt1(i) = sum_other_amt1(i) + de_other_amt1
                 sum_sawo_amt(i) = sum_sawo_amt(i) + de_sawo_amt
				 sum_johab_amt(i) = sum_johab_amt(i) + de_johab_amt
                 sum_hyubjo_amt(i) = sum_hyubjo_amt(i) + de_hyubjo_amt
                 sum_school_amt(i) = sum_school_amt(i) + de_school_amt
                 sum_nhis_bla_amt(i) = sum_nhis_bla_amt(i) + de_nhis_bla_amt
                 sum_long_bla_amt(i) = sum_long_bla_amt(i) + de_long_bla_amt
	             sum_deduct_tot(i) = sum_deduct_tot(i) + de_deduct_tot
				 
				 sum_nps_amt(6) = sum_nps_amt(6) + de_nps_amt
                 sum_nhis_amt(6) = sum_nhis_amt(6) + de_nhis_amt
                 sum_epi_amt(6) = sum_epi_amt(6) + de_epi_amt
	             sum_longcare_amt(6) = sum_longcare_amt(6) + de_longcare_amt
                 sum_income_tax(6) = sum_income_tax(6) + de_income_tax
                 sum_wetax(6) = sum_wetax(6) + de_wetax
	             sum_year_incom_tax(6) = sum_year_incom_tax(6) + de_year_incom_tax
                 sum_year_wetax(6) = sum_year_wetax(6) + de_year_wetax
				 sum_year_incom_tax2(6) = sum_year_incom_tax2(6) + de_year_incom_tax2
                 sum_year_wetax2(6) = sum_year_wetax2(6) + de_year_wetax2
                 sum_other_amt1(6) = sum_other_amt1(6) + de_other_amt1
                 sum_sawo_amt(6) = sum_sawo_amt(6) + de_sawo_amt
				 sum_johab_amt(6) = sum_johab_amt(6) + de_johab_amt
                 sum_hyubjo_amt(6) = sum_hyubjo_amt(6) + de_hyubjo_amt
                 sum_school_amt(6) = sum_school_amt(6) + de_school_amt
                 sum_nhis_bla_amt(6) = sum_nhis_bla_amt(6) + de_nhis_bla_amt
                 sum_long_bla_amt(6) = sum_long_bla_amt(6) + de_long_bla_amt
	             sum_deduct_tot(6) = sum_deduct_tot(6) + de_deduct_tot
	    end if		 
	 next		

	rs.movenext()
loop
rs.close()

    sql = " delete from pay_person_sum " 	
    dbconn.execute(sql)

if view_condi = "전체" then
  sql_sum = " SELECT a.pmg_company,a.pmg_emp_no, saup_count, pmg_base_pay, pmg_meals_pay, pmg_postage_pay, pmg_re_pay, pmg_overtime_pay, " & _
            "   pmg_car_pay, pmg_position_pay, pmg_custom_pay,pmg_job_pay,pmg_job_support,pmg_jisa_pay,pmg_long_pay, " & _
			"   pmg_disabled_pay,pmg_give_total, " & _
			"   de_nps_amt,de_nhis_amt,de_epi_amt,de_longcare_amt,de_income_tax,de_wetax,de_year_incom_tax,de_year_wetax, " & _
			"   de_year_incom_tax2,de_year_wetax2, " & _
			"   de_sawo_amt,de_johab_amt,de_hyubjo_amt,de_school_amt,de_other_amt1,de_nhis_bla_amt,de_long_bla_amt,de_deduct_total " & _
			"   FROM ( " & _
			" select pmg_company,pmg_emp_no,count(*) as saup_count,sum(pmg_base_pay) as pmg_base_pay,sum(pmg_meals_pay) as pmg_meals_pay," & _
            "   sum(pmg_postage_pay) as pmg_postage_pay,sum(pmg_re_pay) as pmg_re_pay,sum(pmg_overtime_pay) as pmg_overtime_pay," & _
			"   sum(pmg_car_pay) as pmg_car_pay,sum(pmg_position_pay) as pmg_position_pay,sum(pmg_custom_pay) as pmg_custom_pay," & _
			"   sum(pmg_job_pay) as pmg_job_pay,sum(pmg_job_support) as pmg_job_support,sum(pmg_jisa_pay) as pmg_jisa_pay," & _
			"   sum(pmg_long_pay) as pmg_long_pay,sum(pmg_disabled_pay) as pmg_disabled_pay,sum(pmg_give_total) as pmg_give_total " & _
			"   from pay_month_give " & _
			"   WHERE (pmg_yymm >= '"+f_yymm+"' and pmg_yymm <= '"+t_yymm+"') and (pmg_id = '1') " & _
			"   group by pmg_company,pmg_emp_no " & _
			"   order by pmg_company,pmg_emp_no ASC " & _
			"   ) a, " & _
			" ( select de_company,de_emp_no,sum(de_nps_amt) as de_nps_amt,sum(de_nhis_amt) as de_nhis_amt,sum(de_epi_amt) as de_epi_amt," & _
            "   sum(de_longcare_amt) as de_longcare_amt,sum(de_income_tax) as de_income_tax,sum(de_wetax) as de_wetax," & _
			"   sum(de_year_incom_tax) as de_year_incom_tax,sum(de_year_wetax) as de_year_wetax," & _
			"   sum(de_year_incom_tax2) as de_year_incom_tax2,sum(de_year_wetax2) as de_year_wetax2,sum(de_sawo_amt) as de_sawo_amt," & _
			"   sum(de_johab_amt) as de_johab_amt,sum(de_hyubjo_amt) as de_hyubjo_amt,sum(de_school_amt) as de_school_amt," & _
			"   sum(de_other_amt1) as de_other_amt1,sum(de_nhis_bla_amt) as de_nhis_bla_amt,sum(de_long_bla_amt) as de_long_bla_amt," & _
			"   sum(de_deduct_total) as de_deduct_total " & _
			"   from pay_month_deduct " & _
			"   WHERE (de_yymm >= '"+f_yymm+"' and de_yymm <= '"+t_yymm+"') and (de_id = '1') " & _	
			"   group by de_company,de_emp_no " & _
			"   order by de_company,de_emp_no ASC " & _
			"   ) b " & _		
			"  WHERE a.pmg_company = b.de_company and a.pmg_emp_no = b.de_emp_no " & _
			"  ORDER BY a.pmg_company,a.pmg_emp_no ASC " 
 else
  sql_sum = " SELECT a.pmg_company,a.pmg_emp_no, saup_count, pmg_base_pay, pmg_meals_pay, pmg_postage_pay, pmg_re_pay, pmg_overtime_pay, " & _
            "   pmg_car_pay, pmg_position_pay, pmg_custom_pay,pmg_job_pay,pmg_job_support,pmg_jisa_pay,pmg_long_pay, " & _
			"   pmg_disabled_pay,pmg_give_total, " & _
			"   de_nps_amt,de_nhis_amt,de_epi_amt,de_longcare_amt,de_income_tax,de_wetax,de_year_incom_tax,de_year_wetax, " & _
			"   de_year_incom_tax2,de_year_wetax2, " & _
			"   de_sawo_amt,de_johab_amt,de_hyubjo_amt,de_school_amt,de_other_amt1,de_nhis_bla_amt,de_long_bla_amt,de_deduct_total " & _
			"   FROM ( " & _
			" select pmg_company,pmg_emp_no,count(*) as saup_count,sum(pmg_base_pay) as pmg_base_pay,sum(pmg_meals_pay) as pmg_meals_pay," & _
            "   sum(pmg_postage_pay) as pmg_postage_pay,sum(pmg_re_pay) as pmg_re_pay,sum(pmg_overtime_pay) as pmg_overtime_pay," & _
			"   sum(pmg_car_pay) as pmg_car_pay,sum(pmg_position_pay) as pmg_position_pay,sum(pmg_custom_pay) as pmg_custom_pay," & _
			"   sum(pmg_job_pay) as pmg_job_pay,sum(pmg_job_support) as pmg_job_support,sum(pmg_jisa_pay) as pmg_jisa_pay," & _
			"   sum(pmg_long_pay) as pmg_long_pay,sum(pmg_disabled_pay) as pmg_disabled_pay,sum(pmg_give_total) as pmg_give_total " & _
			"   from pay_month_give " & _
			"   WHERE (pmg_yymm >= '"+f_yymm+"' and pmg_yymm <= '"+t_yymm+"') and (pmg_id = '1') and (pmg_company = '"+view_condi+"') " & _
			"   group by pmg_company,pmg_emp_no " & _
			"   order by pmg_company,pmg_emp_no ASC " & _
			"   ) a, " & _
			" ( select de_company,de_emp_no,sum(de_nps_amt) as de_nps_amt,sum(de_nhis_amt) as de_nhis_amt,sum(de_epi_amt) as de_epi_amt," & _
            "   sum(de_longcare_amt) as de_longcare_amt,sum(de_income_tax) as de_income_tax,sum(de_wetax) as de_wetax," & _
			"   sum(de_year_incom_tax) as de_year_incom_tax,sum(de_year_wetax) as de_year_wetax," & _
			"   sum(de_year_incom_tax2) as de_year_incom_tax2,sum(de_year_wetax2) as de_year_wetax2,sum(de_sawo_amt) as de_sawo_amt," & _
			"   sum(de_johab_amt) as de_johab_amt,sum(de_hyubjo_amt) as de_hyubjo_amt,sum(de_school_amt) as de_school_amt," & _
			"   sum(de_other_amt1) as de_other_amt1,sum(de_nhis_bla_amt) as de_nhis_bla_amt,sum(de_long_bla_amt) as de_long_bla_amt," & _
			"   sum(de_deduct_total) as de_deduct_total " & _
			"   from pay_month_deduct " & _
			"   WHERE (de_yymm >= '"+f_yymm+"' and de_yymm <= '"+t_yymm+"') and (de_id = '1') and (de_company = '"+view_condi+"') " & _	
			"   group by de_company,de_emp_no " & _
			"   order by de_company,de_emp_no ASC " & _
			"   ) b " & _		
			"  WHERE a.pmg_company = b.de_company and a.pmg_emp_no = b.de_emp_no " & _
			"  ORDER BY a.pmg_company,a.pmg_emp_no ASC " 
end if

rs.Open sql_sum, Dbconn, 1
do until rs.eof

    sql="insert into pay_person_sum (ps_company,ps_emp_no,ps_base_pay,ps_meals_pay,ps_postage_pay,ps_re_pay,ps_overtime_pay,ps_car_pay,ps_position_pay,ps_custom_pay,ps_job_pay,ps_job_support,ps_jisa_pay,ps_long_pay,ps_disabled_pay,ps_give_total,ps_nps_amt,ps_nhis_amt,ps_epi_amt,ps_longcare_amt,ps_income_tax,ps_wetax,ps_year_incom_tax,ps_year_wetax,ps_year_incom_tax2,ps_year_wetax2,ps_sawo_amt,ps_johab_amt,ps_hyubjo_amt,ps_school_amt,ps_other_amt1,ps_nhis_bla_amt,ps_long_bla_amt,ps_deduct_total) values ('"&rs("pmg_company")&"','"&rs("pmg_emp_no")&"','"&rs("pmg_base_pay")&"','"&rs("pmg_meals_pay")&"','"&rs("pmg_postage_pay")&"','"&rs("pmg_re_pay")&"','"&rs("pmg_overtime_pay")&"','"&rs("pmg_car_pay")&"','"&rs("pmg_position_pay")&"','"&rs("pmg_custom_pay")&"','"&rs("pmg_job_pay")&"','"&rs("pmg_job_support")&"','"&rs("pmg_jisa_pay")&"','"&rs("pmg_long_pay")&"','"&rs("pmg_disabled_pay")&"','"&rs("pmg_give_total")&"','"&rs("de_nps_amt")&"','"&rs("de_nhis_amt")&"','"&rs("de_epi_amt")&"','"&rs("de_longcare_amt")&"','"&rs("de_income_tax")&"','"&rs("de_wetax")&"','"&rs("de_year_incom_tax")&"','"&rs("de_year_wetax")&"','"&rs("de_year_incom_tax2")&"','"&rs("de_year_wetax2")&"','"&rs("de_sawo_amt")&"','"&rs("de_johab_amt")&"','"&rs("de_hyubjo_amt")&"','"&rs("de_school_amt")&"','"&rs("de_other_amt1")&"','"&rs("de_nhis_bla_amt")&"','"&rs("de_long_bla_amt")&"','"&rs("de_deduct_total")&"')"
	
	dbconn.execute(sql)

	rs.movenext()
loop
rs.close()				   

Sql = "select * from pay_person_sum"
Rs.Open Sql, Dbconn, 1
if not Rs.eof then
   do until Rs.eof

    ps_emp_no = rs("ps_emp_no")
	ps_company = rs("ps_company")
	
	Sql = "select * from emp_master where emp_no = '"+ps_emp_no+"'"
	Set Rs_emp = DbConn.Execute(SQL)
	if not Rs_emp.EOF or not Rs_emp.BOF then
	        
			emp_name = rs_emp("emp_name")
			emp_grade = rs_emp("emp_grade")
			emp_bonbu = rs_emp("emp_bonbu")
			emp_saupbu = rs_emp("emp_saupbu")
			emp_org_name = rs_emp("emp_org_name")
			cost_center = rs_emp("cost_center")
			cost_group = rs_emp("cost_group")
			
			sql="select max(cost_group) as max_cost_group,max(cost_center) as max_cost_center,max(pmg_bonbu) as max_pmg_bonbu,max(pmg_saupbu) as max_pmg_saupbu,max(pmg_org_name) as max_pmg_org_name from pay_month_give where (pmg_yymm >= '"+f_yymm+"' and pmg_yymm <= '"+t_yymm+"') and (pmg_id = '1') and (pmg_company = '"+ps_company+"') and (pmg_emp_no = '"+ps_emp_no+"')"
	        set rs_max=dbconn.execute(sql)
			
			if	isnull(rs_max("max_cost_group"))  then
			        cost_group = rs_emp("cost_group")
		        else
					cost_center = rs_max("max_cost_center")
			        cost_group = rs_max("max_cost_group")
					emp_bonbu = rs_max("max_pmg_bonbu")
			        emp_saupbu = rs_max("max_pmg_saupbu")
			        emp_org_name = rs_max("max_pmg_org_name")
					if rs_max("max_cost_center") = "" then
					       cost_center = rs_emp("cost_center")
				    end if
		    end if
	   
	        sql = "update pay_person_sum set ps_emp_name='"&emp_name&"',ps_grade='"&emp_grade&"',ps_bonbu='"&emp_bonbu&"',ps_saupbu='"&emp_saupbu&"',ps_org_name='"&emp_org_name&"',ps_cost_group='"&cost_group&"',ps_cost_center='"&cost_center&"' where (ps_company = '"+ps_company+"' ) and (ps_emp_no = '"+ps_emp_no+"')"
		
		   dbconn.execute(sql)	
		else
		    Sql = "select * from pay_data_conver where company = '"+ps_company+"' and emp_no = '"+ps_emp_no+"'"
	        Set Rs_old = DbConn.Execute(SQL)
	        if not Rs_old.EOF or not Rs_old.BOF then
	        
			       emp_name = Rs_old("emp_name")
			       emp_grade = Rs_old("emp_grade")
			       emp_bonbu = ""
			       emp_saupbu = ""
		           emp_org_name = Rs_old("emp_org_name")
			       cost_center = Rs_old("cost_center")
			       cost_group = Rs_old("cost_group")
	   
	               sql = "update pay_person_sum set ps_emp_name='"&emp_name&"',ps_grade='"&emp_grade&"',ps_bonbu='"&emp_bonbu&"',ps_saupbu='"&emp_saupbu&"',ps_org_name='"&emp_org_name&"',ps_cost_group='"&cost_group&"',ps_cost_center='"&cost_center&"' where (ps_company = '"+ps_company+"' ) and (ps_emp_no = '"+ps_emp_no+"')"
		
		           dbconn.execute(sql)	
		     end if
			 Rs_old.close()	
	end if	 
	    Rs_emp.close()	
	    Rs.MoveNext()
  loop		
end if
rs.close()

Sql = "select * from pay_person_sum ORDER BY ps_cost_center,ps_cost_group,ps_saupbu,ps_org_name,ps_emp_no ASC" 
Rs.Open Sql, Dbconn, 1

curr_yyyy = mid(cstr(from_yymm),1,4)
curr_mm = mid(cstr(from_yymm),5,2)

title_line = cstr(f_yymm) + " ∼ " + cstr(t_yymm) + "월 " + " 개인별 급여현황(기간별)-" + view_condi 

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
			function getPageCode(){
				return "7 1";
			}
		</script>
		<script type="text/javascript">
		    $(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});	  

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
			function scrollAll() {
			//  document.all.leftDisplay2.scrollTop = document.all.mainDisplay2.scrollTop;
			  document.all.topLine2.scrollLeft = document.all.mainDisplay2.scrollLeft;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_report_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_person_term_report.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>◈ 검색◈</dt>
                        <dd>
                            <p>
                             <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where  org_level = '회사' ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:130px">
                                    <option value="전체" <%If view_condi = "전체" then %>selected<% end if %>>전체</option>
                			  <% 
								do until rs_org.eof 
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()  
								loop 
								rs_org.Close()
							  %>
            					</select>
                                </label>
                                <label>
								<strong>귀속년월(시작월) : </strong>
                                    <select name="from_yymm" id="from_yymm" type="text" value="<%=f_yymm%>" style="width:90px">
                                    <%	for i = 100 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If f_yymm = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								</label>
                                <label>
								<strong> ∼ 종료월 : </strong>
                                    <select name="to_yymm" id="to_yymm" type="text" value="<%=t_yymm%>" style="width:90px">
                                    <%	for i = 100 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If t_yymm = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
                <table cellpadding="0" cellspacing="0">
				  <tr>
                   	<td>
      				<DIV id="topLine2" style="width:1200px;overflow:hidden;">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
                            <col width="7%" >
                            
                            <col width="7%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
							<col width="7%" >
                            
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="*" >
                            <col width="2%" >
						</colgroup>
						<thead>
							<tr>
				               <th rowspan="2" class="first" scope="col">성명</th>
                               <th rowspan="2" scope="col">부서/직급</th>
				               <th colspan="6" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;">지급항목</th>
                               <th colspan="7" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;">공제항목</th>
                               <th rowspan="2" scope="col">지급액</th>
                               <th rowspan="2" scope="col">비고</th>
			                </tr>
                            <tr>
								<th scope="col" style=" border-left:1px solid #e3e3e3;">기본급</th>
								<th scope="col">식대</th>
								<th scope="col">통신비</th>
								<th scope="col">연장</th>
								<th scope="col">소급등</th>
								<th scope="col">지급액계</th>
								<th scope="col">국민<br>연금</th>
								<th scope="col">건강<br>보험</th>
								<th scope="col">고용<br>보험</th>
                                <th scope="col">장기<br>요양</th>
                                <th scope="col">사우회<br>비 등</th>
                                <th scope="col">소득세등</th>
                                <th scope="col">공제액계</th>
							</tr>
						</thead>
                    </table>
                    </DIV>
					</td>
                  </tr>
                  <tr>
                    <td valign="top">
				    <DIV id="mainDisplay2" style="width:1200;height:400px;overflow:scroll" onscroll="scrollAll()">
					<table cellpadding="0" cellspacing="0" class="scrollList">                            
						<colgroup>
							<col width="5%" >
                            <col width="7%" >
                            
                            <col width="7%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
							<col width="7%" >
                            
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="*" >
                            <col width="2%" >
						</colgroup>
                        <tbody>
						<%
						do until rs.eof
							  ps_emp_no = rs("ps_emp_no")
							  ps_give_total = rs("ps_give_total")
							  
							  pmg_curr_pay = int(rs("ps_give_total")) + int(rs("ps_deduct_total"))
							  
							  hap_give_hap = int(rs("ps_re_pay")) + int(rs("ps_car_pay")) + int(rs("ps_position_pay")) + int(rs("ps_custom_pay")) + int(rs("ps_job_pay")) + int(rs("ps_job_support")) + int(rs("ps_jisa_pay")) + int(rs("ps_long_pay")) + int(rs("ps_disabled_pay"))
							  
							  hap_de_other = int(rs("ps_sawo_amt")) + int(rs("ps_johab_amt")) + int(rs("ps_hyubjo_amt")) + int(rs("ps_school_amt")) + int(rs("ps_other_amt1")) + int(rs("ps_nhis_bla_amt")) + int(rs("ps_long_bla_amt"))
							  
							  hap_de_tax = int(rs("ps_income_tax")) + int(rs("ps_wetax")) + int(rs("ps_year_incom_tax")) + int(rs("ps_year_wetax")) + int(rs("ps_year_incom_tax2")) + int(rs("ps_year_wetax2"))
			   
	           			%>
							<tr>
								<td class="first" style="font-size:11px;"><%=rs("ps_emp_name")%><br>(<%=rs("ps_emp_no")%>)</td>
                                <td style="font-size:11px;"><%=rs("ps_org_name")%><br>(<%=rs("ps_grade")%>)</td>
                                <td class="right"><%=formatnumber(rs("ps_base_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("ps_meals_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("ps_postage_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("ps_overtime_pay"),0)%></td>
                                <td class="right"><%=formatnumber(hap_give_hap,0)%></td>
                                <td class="right"><%=formatnumber(rs("ps_give_total"),0)%></td>
                                
                                <td class="right"><%=formatnumber(rs("ps_nps_amt"),0)%></td>
                                <td class="right"><%=formatnumber(rs("ps_nhis_amt"),0)%></td>
                                <td class="right"><%=formatnumber(rs("ps_epi_amt"),0)%></td>
                                <td class="right"><%=formatnumber(rs("ps_longcare_amt"),0)%></td>
                                <td class="right"><%=formatnumber(hap_de_other,0)%></td>
                                <td class="right"><%=formatnumber(hap_de_tax,0)%></td>
                                <td class="right"><%=formatnumber(rs("ps_deduct_total"),0)%></td>

                                <td class="right"><%=formatnumber(pmg_curr_pay,0)%></td>
                                <td class="right">&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()

                              sum_curr_pay = sum_give_tot(6) + sum_deduct_tot(6)
							  
							  hap_give_hap = sum_re_pay(6) + sum_car_pay(6) + sum_position_pay(6) + sum_custom_pay(6) + sum_job_pay(6) + sum_job_support(6) + sum_jisa_pay(6) + sum_long_pay(6) + sum_disabled_pay(6)
							  
							  hap_de_other = sum_sawo_amt(6) + sum_johab_amt(6) + sum_hyubjo_amt(6) + sum_school_amt(6) + sum_other_amt1(6) + sum_nhis_bla_amt(6) + sum_long_bla_amt(6)
							  
							  hap_de_tax = sum_income_tax(6) + sum_wetax(6) + sum_year_incom_tax(6) + sum_year_wetax(6) + sum_year_incom_tax2(6) + sum_year_wetax2(6)

						%>
                          	<tr>
                                <td colspan="2" class="first" style="background:#ffe8e8;">총계(<%=formatnumber(pay_count(6),0)%>&nbsp;명)</td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_base_pay(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_meals_pay(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_postage_pay(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_overtime_pay(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(hap_give_hap,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_give_tot(6),0)%></td>
                                
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_nps_amt(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_nhis_amt(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_epi_amt(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_longcare_amt(6),0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(hap_de_other,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(hap_de_tax,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_deduct_tot(6),0)%></td>
                                
                                <td class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(sum_curr_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#ffe8e8;">&nbsp;</td>
							</tr>
                         <%
						    for i = 1 to 6 
                        	     if	com_tab(i) <> "" then
								 
								 sum_curr_pay = sum_give_tot(i) + sum_deduct_tot(i)
							  
							     hap_give_hap = sum_re_pay(i) + sum_car_pay(i) + sum_position_pay(i) + sum_custom_pay(i) + sum_job_pay(i) + sum_job_support(i) + sum_jisa_pay(i) + sum_long_pay(i) + sum_disabled_pay(i)
							  
							     hap_de_other = sum_sawo_amt(i) + sum_johab_amt(i) + sum_hyubjo_amt(i) + sum_school_amt(i) + sum_other_amt1(i) + sum_nhis_bla_amt(i) + sum_long_bla_amt(i)
							  
							     hap_de_tax = sum_income_tax(i) + sum_wetax(i) + sum_year_incom_tax(i) + sum_year_wetax(i) + sum_year_incom_tax2(i) + sum_year_wetax2(i)
						 %>	
                            <tr>
                                <td colspan="2" class="first" style="background:#eeffff;"><%=com_tab(i)%>(<%=formatnumber(pay_count(i),0)%>&nbsp;명)</td>
                                <td class="right" style="font-size:11px; background:#eeffff;"><%=formatnumber(sum_base_pay(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#eeffff;"><%=formatnumber(sum_meals_pay(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#eeffff;"><%=formatnumber(sum_postage_pay(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#eeffff;"><%=formatnumber(sum_overtime_pay(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#eeffff;"><%=formatnumber(hap_give_hap,0)%></td>
                                <td class="right" style="font-size:11px; background:#eeffff;"><%=formatnumber(sum_give_tot(i),0)%></td>
                                
                                <td class="right" style="font-size:11px; background:#eeffff;"><%=formatnumber(sum_nps_amt(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#eeffff;"><%=formatnumber(sum_nhis_amt(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#eeffff;"><%=formatnumber(sum_epi_amt(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#eeffff;"><%=formatnumber(sum_longcare_amt(i),0)%></td>
                                <td class="right" style="font-size:11px; background:#eeffff;"><%=formatnumber(hap_de_other,0)%></td>
                                <td class="right" style="font-size:11px; background:#eeffff;"><%=formatnumber(hap_de_tax,0)%></td>
                                <td class="right" style="font-size:11px; background:#eeffff;"><%=formatnumber(sum_deduct_tot(i),0)%></td>
                                
                                <td class="right" style="font-size:11px; background:#eeffff;"><%=formatnumber(sum_curr_pay,0)%></td>
                                <td class="right" style="font-size:11px; background:#eeffff;">&nbsp;</td>
							</tr>
                         <%
							     end if
						    next
					     %>                            
						</tbody>
					</table>
                    </DIV>
					</td>
                  </tr>
				</table>
                <br>                        
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                    <td width="25%">
					<div class="btnleft">
                    <a href="insa_pay_person_term_excel.asp?view_condi=<%=view_condi%>&from_yymm=<%=f_yymm%>&to_yymm=<%=t_yymm%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                    </td>
                    <td width="50%">
                    </td>
				    <td width="25%">
					<div class="btnRight">
                    
					</div>                  
                    </td>        
                    </td> 
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

