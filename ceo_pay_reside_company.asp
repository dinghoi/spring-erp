<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)
dim quarter_tab(8,2)
dim year_tab(3,2)

dim com_tab(5)
dim pay_count(5)
dim sum_base_pay(5)
dim sum_meals_pay(5)
dim sum_postage_pay(5)
dim sum_re_pay(5)
dim sum_overtime_pay(5)
dim sum_car_pay(5)
dim sum_position_pay(5)
dim sum_custom_pay(5)
dim sum_job_pay(5)
dim sum_job_support(5)
dim sum_jisa_pay(5)
dim sum_long_pay(5)
dim sum_disabled_pay(5)
dim sum_give_tot(5)

dim sum_nps_amt(5)
dim sum_nhis_amt(5)
dim sum_epi_amt(5)
dim sum_longcare_amt(5)
dim sum_income_tax(5)
dim sum_wetax(5)
dim sum_year_incom_tax(5)
dim sum_year_wetax(5)
dim sum_year_incom_tax2(5)
dim sum_year_wetax2(5)
dim sum_other_amt1(5)
dim sum_sawo_amt(5)
dim sum_hyubjo_amt(5)
dim sum_school_amt(5)
dim sum_nhis_bla_amt(5)
dim sum_long_bla_amt(5)
dim sum_deduct_tot(5)

be_pg = "ceo_pay_reside_company.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

Page=Request("page")
view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	pmg_yymm=Request.form("pmg_yymm")
  else
	view_condi = request("view_condi")
	pmg_yymm=request("pmg_yymm")
end if

if view_condi = "" then
	view_condi = "케이원정보통신"
	curr_dd = cstr(datepart("d",now))
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	'pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	pmg_yymm = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
	
	for i = 1 to 5
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
        sum_hyubjo_amt(i) = 0
        sum_school_amt(i) = 0
        sum_nhis_bla_amt(i) = 0
        sum_long_bla_amt(i) = 0
        sum_deduct_tot(i) = 0
    next
	
	sum_curr_pay = 0	
	
end if

' 최근3개년도 테이블로 생성
year_tab(3,1) = mid(now(),1,4)
year_tab(3,2) = cstr(year_tab(3,1)) + "년"
year_tab(2,1) = cint(mid(now(),1,4)) - 1
year_tab(2,2) = cstr(year_tab(2,1)) + "년"
year_tab(1,1) = cint(mid(now(),1,4)) - 2
year_tab(1,2) = cstr(year_tab(1,1)) + "년"

' 분기 테이블 생성
curr_mm = mid(now(),6,2)
if curr_mm > 0 and curr_mm < 4 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "1"
end if
if curr_mm > 3 and curr_mm < 7 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "2"
end if
if curr_mm > 6 and curr_mm < 10 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "3"
end if
if curr_mm > 9 and curr_mm < 13 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "4"
end if

quarter_tab(8,2) = cstr(mid(quarter_tab(8,1),1,4)) + "년 " + cstr(mid(quarter_tab(8,1),5,1)) + "/4분기"

for i = 7 to 1 step -1
	cal_quarter = cint(quarter_tab(i+1,1)) - 1
	if cstr(mid(cal_quarter,5,1)) = "0" then
		quarter_tab(i,1) = cstr(cint(mid(cal_quarter,1,4))-1) + "4"
	  else
		quarter_tab(i,1) = cal_quarter
	end if	 
	quarter_tab(i,2) = cstr(mid(quarter_tab(i,1),1,4)) + "년 " + cstr(mid(quarter_tab(i,1),5,1)) + "/4분기"
next

' 년월 테이블생성
cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
'cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
month_tab(24,1) = cal_month
view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
month_tab(24,2) = view_month
for i = 1 to 23
	cal_month = cstr(int(cal_month) - 1)
	if mid(cal_month,5) = "00" then
		cal_year = cstr(int(mid(cal_month,1,4)) - 1)
		cal_month = cal_year + "12"
	end if	 
	view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
	j = 24 - i
	month_tab(j,1) = cal_month
	month_tab(j,2) = view_month
next

pgsize = 10 ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
if view_condi = "전체" then
      com_tab(1) = "케이원정보통신"
	  com_tab(2) = "휴디스"
	  com_tab(3) = "케이네트웍스"
	  com_tab(4) = "에스유에이치"
	  com_tab(5) = "합계"
	  where_sql = " WHERE (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1')" 
   else  
      com_tab(1) = view_condi
	  com_tab(5) = "합계"
	  where_sql = " WHERE (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"')"
end if   

sql = "select * from pay_month_give " + where_sql + order_sql

'Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    emp_no = rs("pmg_emp_no")
	pmg_company = rs("pmg_company")
				  
    for i = 1 to 5
        if com_tab(i) = rs("pmg_company") then
	             pay_count(i) = pay_count(i) + 1
				 pay_count(5) = pay_count(5) + 1
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
				 
				 sum_base_pay(5) = sum_base_pay(5) + int(rs("pmg_base_pay"))
                 sum_meals_pay(5) = sum_meals_pay(5) + int(rs("pmg_meals_pay"))
                 sum_postage_pay(5) = sum_postage_pay(5) + int(rs("pmg_postage_pay"))
                 sum_re_pay(5) = sum_re_pay(5) + int(rs("pmg_re_pay"))
                 sum_overtime_pay(5) = sum_overtime_pay(5) + int(rs("pmg_overtime_pay"))
                 sum_car_pay(5) = sum_car_pay(5) + int(rs("pmg_car_pay"))
                 sum_position_pay(5) = sum_position_pay(5) + int(rs("pmg_position_pay"))
                 sum_custom_pay(5) = sum_custom_pay(5) + int(rs("pmg_custom_pay"))
                 sum_job_pay(5) = sum_job_pay(5) + int(rs("pmg_job_pay"))
                 sum_job_support(5) = sum_job_support(5) + int(rs("pmg_job_support"))
                 sum_jisa_pay(5) = sum_jisa_pay(5) + int(rs("pmg_jisa_pay"))
                 sum_long_pay(5) = sum_long_pay(5) + int(rs("pmg_long_pay"))
                 sum_disabled_pay(5) = sum_disabled_pay(5) + int(rs("pmg_disabled_pay"))
                 sum_give_tot(5) = sum_give_tot(5) + int(rs("pmg_give_total"))
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
            de_hyubjo_amt = 0
            de_school_amt = 0
            de_nhis_bla_amt = 0
            de_long_bla_amt = 0
		    de_deduct_tot = 0
     end if
     Rs_dct.close()
     for i = 1 to 5
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
                 sum_hyubjo_amt(i) = sum_hyubjo_amt(i) + de_hyubjo_amt
                 sum_school_amt(i) = sum_school_amt(i) + de_school_amt
                 sum_nhis_bla_amt(i) = sum_nhis_bla_amt(i) + de_nhis_bla_amt
                 sum_long_bla_amt(i) = sum_long_bla_amt(i) + de_long_bla_amt
	             sum_deduct_tot(i) = sum_deduct_tot(i) + de_deduct_tot
				 
				 sum_nps_amt(5) = sum_nps_amt(5) + de_nps_amt
                 sum_nhis_amt(5) = sum_nhis_amt(5) + de_nhis_amt
                 sum_epi_amt(5) = sum_epi_amt(5) + de_epi_amt
	             sum_longcare_amt(5) = sum_longcare_amt(5) + de_longcare_amt
                 sum_income_tax(5) = sum_income_tax(5) + de_income_tax
                 sum_wetax(5) = sum_wetax(5) + de_wetax
	             sum_year_incom_tax(5) = sum_year_incom_tax(5) + de_year_incom_tax
                 sum_year_wetax(5) = sum_year_wetax(5) + de_year_wetax
				 sum_year_incom_tax2(5) = sum_year_incom_tax2(5) + de_year_incom_tax2
                 sum_year_wetax2(5) = sum_year_wetax2(5) + de_year_wetax2
                 sum_other_amt1(5) = sum_other_amt1(5) + de_other_amt1
                 sum_sawo_amt(5) = sum_sawo_amt(5) + de_sawo_amt
                 sum_hyubjo_amt(5) = sum_hyubjo_amt(5) + de_hyubjo_amt
                 sum_school_amt(5) = sum_school_amt(5) + de_school_amt
                 sum_nhis_bla_amt(5) = sum_nhis_bla_amt(5) + de_nhis_bla_amt
                 sum_long_bla_amt(5) = sum_long_bla_amt(5) + de_long_bla_amt
	             sum_deduct_tot(5) = sum_deduct_tot(5) + de_deduct_tot
	    end if		 
	 next				

	rs.movenext()
loop
rs.close()

if view_condi = "전체" then
      Sql = " SELECT a.cost_group, saup_count, pmg_base_pay, pmg_meals_pay, pmg_postage_pay, pmg_re_pay, pmg_overtime_pay, " & _
            "   pmg_car_pay, pmg_position_pay, pmg_custom_pay,pmg_job_pay,pmg_job_support,pmg_jisa_pay,pmg_long_pay, " & _
			"   pmg_disabled_pay,pmg_give_total, " & _
			"   de_nps_amt,de_nhis_amt,de_epi_amt,de_longcare_amt,de_income_tax,de_wetax,de_year_incom_tax,de_year_wetax, " & _
			"   de_year_incom_tax2,de_year_wetax2, " & _
			"   de_sawo_amt,de_johab_amt,de_hyubjo_amt,de_school_amt,de_other_amt1,de_nhis_bla_amt,de_long_bla_amt,de_deduct_total " & _
			"   FROM ( " & _
			" select cost_group,count(*) as saup_count,sum(pmg_base_pay) as pmg_base_pay,sum(pmg_meals_pay) as pmg_meals_pay," & _
            "   sum(pmg_postage_pay) as pmg_postage_pay,sum(pmg_re_pay) as pmg_re_pay,sum(pmg_overtime_pay) as pmg_overtime_pay," & _
			"   sum(pmg_car_pay) as pmg_car_pay,sum(pmg_position_pay) as pmg_position_pay,sum(pmg_custom_pay) as pmg_custom_pay," & _
			"   sum(pmg_job_pay) as pmg_job_pay,sum(pmg_job_support) as pmg_job_support,sum(pmg_jisa_pay) as pmg_jisa_pay," & _
			"   sum(pmg_long_pay) as pmg_long_pay,sum(pmg_disabled_pay) as pmg_disabled_pay,sum(pmg_give_total) as pmg_give_total " & _
			"   from pay_month_give " & _
			"   where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') group by cost_group " & _
			"   order by cost_group " & _
			"   ) a, " & _
			" ( select cost_group,sum(de_nps_amt) as de_nps_amt,sum(de_nhis_amt) as de_nhis_amt,sum(de_epi_amt) as de_epi_amt," & _
            "   sum(de_longcare_amt) as de_longcare_amt,sum(de_income_tax) as de_income_tax,sum(de_wetax) as de_wetax," & _
			"   sum(de_year_incom_tax) as de_year_incom_tax,sum(de_year_wetax) as de_year_wetax," & _
			"   sum(de_year_incom_tax2) as de_year_incom_tax2,sum(de_year_wetax2) as de_year_wetax2,sum(de_sawo_amt) as de_sawo_amt," & _
			"   sum(de_johab_amt) as de_johab_amt,sum(de_hyubjo_amt) as de_hyubjo_amt,sum(de_school_amt) as de_school_amt," & _
			"   sum(de_other_amt1) as de_other_amt1,sum(de_nhis_bla_amt) as de_nhis_bla_amt,sum(de_long_bla_amt) as de_long_bla_amt," & _
			"   sum(de_deduct_total) as de_deduct_total " & _
			"   from pay_month_deduct " & _
			"   where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') group by cost_group " & _	
			"   order by cost_group " & _
			"   ) b " & _		
			"  WHERE a.cost_group = b.cost_group " & _
			"  ORDER BY a.cost_group ASC " 
    else
      Sql = " SELECT a.cost_group, saup_count, pmg_base_pay, pmg_meals_pay, pmg_postage_pay, pmg_re_pay, pmg_overtime_pay, " & _
            "   pmg_car_pay, pmg_position_pay, pmg_custom_pay,pmg_job_pay,pmg_job_support,pmg_jisa_pay,pmg_long_pay, " & _
			"   pmg_disabled_pay,pmg_give_total, " & _
			"   de_nps_amt,de_nhis_amt,de_epi_amt,de_longcare_amt,de_income_tax,de_wetax,de_year_incom_tax,de_year_wetax, " & _
			"   de_year_incom_tax2,de_year_wetax2, " & _
			"   de_sawo_amt,de_johab_amt,de_hyubjo_amt,de_school_amt,de_other_amt1,de_nhis_bla_amt,de_long_bla_amt,de_deduct_total " & _
			"   FROM ( " & _
			" select cost_group,count(*) as saup_count,sum(pmg_base_pay) as pmg_base_pay,sum(pmg_meals_pay) as pmg_meals_pay," & _
            "   sum(pmg_postage_pay) as pmg_postage_pay,sum(pmg_re_pay) as pmg_re_pay,sum(pmg_overtime_pay) as pmg_overtime_pay," & _
			"   sum(pmg_car_pay) as pmg_car_pay,sum(pmg_position_pay) as pmg_position_pay,sum(pmg_custom_pay) as pmg_custom_pay," & _
			"   sum(pmg_job_pay) as pmg_job_pay,sum(pmg_job_support) as pmg_job_support,sum(pmg_jisa_pay) as pmg_jisa_pay," & _
			"   sum(pmg_long_pay) as pmg_long_pay,sum(pmg_disabled_pay) as pmg_disabled_pay,sum(pmg_give_total) as pmg_give_total " & _
			"   from pay_month_give " & _
			"   where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') group by cost_group " & _
			"   order by cost_group " & _
			"   ) a, " & _
			" ( select cost_group,sum(de_nps_amt) as de_nps_amt,sum(de_nhis_amt) as de_nhis_amt,sum(de_epi_amt) as de_epi_amt," & _
            "   sum(de_longcare_amt) as de_longcare_amt,sum(de_income_tax) as de_income_tax,sum(de_wetax) as de_wetax," & _
			"   sum(de_year_incom_tax) as de_year_incom_tax,sum(de_year_wetax) as de_year_wetax," & _
			"   sum(de_year_incom_tax2) as de_year_incom_tax2,sum(de_year_wetax2) as de_year_wetax2,sum(de_sawo_amt) as de_sawo_amt," & _
			"   sum(de_johab_amt) as de_johab_amt,sum(de_hyubjo_amt) as de_hyubjo_amt,sum(de_school_amt) as de_school_amt," & _
			"   sum(de_other_amt1) as de_other_amt1,sum(de_nhis_bla_amt) as de_nhis_bla_amt,sum(de_long_bla_amt) as de_long_bla_amt," & _
			"   sum(de_deduct_total) as de_deduct_total " & _
			"   from pay_month_deduct " & _
			"   where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_company = '"+view_condi+"') group by cost_group " & _	
			"   order by cost_group " & _
			"   ) b " & _		
			"  WHERE a.cost_group = b.cost_group " & _
			"  ORDER BY a.cost_group ASC " 
end if
Rs.Open Sql, Dbconn, 1

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = cstr(curr_yyyy) + "년 " + cstr(curr_mm) + "월 " + " 상주회사별 급여현황"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>임원 정보 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "4 1";
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
			<!--#include virtual = "/include/ceo_header.asp" -->
            <!--#include virtual = "/include/ceo_pay_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="ceo_pay_reside_company.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>◈ 검색◈</dt>
                        <dd>
                            <p>
                             <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where isNull(org_end_date) and org_level = '회사' ORDER BY org_code ASC"
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
								<strong>귀속년월 : </strong>
                                    <select name="pmg_yymm" id="pmg_yymm" type="text" value="<%=pmg_yymm%>" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If pmg_yymm = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
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
							<col width="*" >
							<col width="5%" >
                            <col width="9%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="9%" >
							<col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
							<col width="9%" > 
                            <col width="9%" >
                            <col width="5%" >
						</colgroup>
						<thead>
							<tr>
				               <th rowspan="2" class="first" scope="col" >상주(처)회사</th>
                               <th rowspan="2" scope="col" >인원</th>
				               <th colspan="5" scope="col" style="background:#FFFFE6;">기본급여 및 제수당</th>
                               <th colspan="4" scope="col" style="background:#E0FFFF;">공제 및 차인지급액</th>
                               <th rowspan="2" scope="col" >지급액</th>
                               <th rowspan="2" scope="col">비고</th>
			                </tr>
                            <tr>
								<td scope="col" style=" border-left:1px solid #e3e3e3; background:#f8f8f8">기본급</td>
								<td scope="col" style="background:#f8f8f8;">식대</td>  
								<td scope="col" style="background:#f8f8f8;">연장근로<br>수당</td>
                                <td scope="col" style="background:#f8f8f8;">통신비 등</td>
                                <td scope="col" style="background:#f8f8f8;">지급소계</td>
								<td scope="col" style="background:#f8f8f8;">4대보험</td>
                                <td scope="col" style="background:#f8f8f8;">소득세 등</td>
								<td scope="col" style="background:#f8f8f8;">기타공제등</td>
                                <td scope="col" style="background:#f8f8f8;">예수금계</td>
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
							<col width="*" >
							<col width="5%" >
                            <col width="9%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="9%" >
							<col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
							<col width="9%" > 
                            <col width="9%" >
                            <col width="5%" >
						</colgroup>                        
						<tbody>
					<%
						do until rs.eof
							  pmg_give_tot = cdbl(rs("pmg_give_total"))
							  
							  sub_give_hap = cdbl(rs("pmg_postage_pay")) + cdbl(rs("pmg_re_pay")) + cdbl(rs("pmg_car_pay")) + cdbl(rs("pmg_position_pay")) + cdbl(rs("pmg_custom_pay")) + cdbl(rs("pmg_job_pay")) + cdbl(rs("pmg_job_support")) + cdbl(rs("pmg_jisa_pay")) + cdbl(rs("pmg_long_pay")) + cdbl(rs("pmg_disabled_pay"))
							
							saupbu_name = rs("cost_group")
							if saupbu_name = "" or saupbu_name = " " or isnull(saupbu_name) then
							    saupbu_name = view_condi
							end if
							  
	           			%>
							<tr>
								<td class="first"><%=saupbu_name%>&nbsp;</td>
                                <td class="first" style=" border-left:1px solid #e3e3e3;"><%=rs("saup_count")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("pmg_base_pay"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("pmg_meals_pay"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("pmg_overtime_pay"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sub_give_hap,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("pmg_give_total"),0)%>&nbsp;</td>
                       <%  
                                  pmg_curr_pay = cdbl(rs("pmg_give_total")) - cdbl(rs("de_deduct_total"))
							  
							      hap_de_insur = cdbl(rs("de_nps_amt")) + cdbl(rs("de_nhis_amt")) + cdbl(rs("de_epi_amt")) + cdbl(rs("de_longcare_amt"))
							      hap_de_tax = cdbl(rs("de_income_tax")) + cdbl(rs("de_wetax")) + cdbl(rs("de_year_incom_tax")) + cdbl(rs("de_year_wetax")) + cdbl(rs("de_year_incom_tax2")) + cdbl(rs("de_year_wetax2"))
							      hap_de_other = cdbl(rs("de_other_amt1")) + cdbl(rs("de_sawo_amt")) + cdbl(rs("de_hyubjo_amt")) + cdbl(rs("de_school_amt")) + cdbl(rs("de_nhis_bla_amt")) + cdbl(rs("de_long_bla_amt"))
								  hap_deduct_tot = hap_de_insur + hap_de_tax + hap_de_other
                       %>
                                <td class="right"><%=formatnumber(hap_de_insur,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(hap_de_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(hap_de_other,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(hap_deduct_tot,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(pmg_curr_pay,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
					   <%
							rs.movenext()
						loop
						rs.close()
						
						sum_curr_pay = sum_give_tot(5) - sum_deduct_tot(5)
						
						sum_give_hap = sum_postage_pay(5) + sum_re_pay(5) + sum_car_pay(5) + sum_position_pay(5) + sum_custom_pay(5) + sum_job_pay(5) + sum_job_support(5) + sum_jisa_pay(5) + sum_long_pay(5) + sum_disabled_pay(5)
						sum_de_insur =sum_nps_amt(5) + sum_nhis_amt(5) + sum_epi_amt(5) + sum_longcare_amt(5)
						sum_de_tax =sum_income_tax(5) + sum_wetax(5) + sum_year_incom_tax(5) + sum_year_wetax(5) + sum_year_incom_tax2(5) + sum_year_wetax2(5)
						sum_de_other =sum_other_amt1(5) + sum_sawo_amt(5) + sum_hyubjo_amt(5) + sum_school_amt(5) + sum_nhis_bla_amt(5) + sum_long_bla_amt(5)
						%>
                          	<tr>
                                <th class="first">총계</th>
                                <th class="right"><%=formatnumber(pay_count(5),0)%>&nbsp;명</th>
                                <th class="right"><%=formatnumber(sum_base_pay(5),0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_meals_pay(5),0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_overtime_pay(5),0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_give_hap,0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_give_tot(5),0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_de_insur,0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_de_tax,0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_de_other,0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_deduct_tot(5),0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(sum_curr_pay,0)%>&nbsp;</th>
                                <th class="right">&nbsp;</th>
							</tr>
                         <%
						    for i = 1 to 5 
                        	     if	com_tab(i) <> "" then
								 
								    sum_curr_pay = sum_give_tot(i) - sum_deduct_tot(i)
						
						            sum_give_hap = sum_postage_pay(i) + sum_re_pay(i) + sum_car_pay(i) + sum_position_pay(i) + sum_custom_pay(i) + sum_job_pay(i) + sum_job_support(i) + sum_jisa_pay(i) + sum_long_pay(i) + sum_disabled_pay(i)
						            sum_de_insur =sum_nps_amt(i) + sum_nhis_amt(i) + sum_epi_amt(i) + sum_longcare_amt(i)
						            sum_de_tax =sum_income_tax(i) + sum_wetax(i) + sum_year_incom_tax(i) + sum_year_wetax(i) + sum_year_incom_tax2(i) + sum_year_wetax2(i)
						            sum_de_other =sum_other_amt1(i) + sum_sawo_amt(i) + sum_hyubjo_amt(i) + sum_school_amt(i) + sum_nhis_bla_amt(i) + sum_long_bla_amt(i)
						 %>	
                            <tr>
                                <td class="first"><%=com_tab(i)%></td>
                                <td class="right"><%=formatnumber(pay_count(i),0)%>&nbsp;명</td>
                                <td class="right"><%=formatnumber(sum_base_pay(i),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_meals_pay(i),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_overtime_pay(i),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_give_hap,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_give_tot(i),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_de_insur,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_de_tax,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_de_other,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_deduct_tot(i),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_curr_pay,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
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
                    </td>
                    <td width="50%">
                    </td>
				    <td width="25%">
                    </td> 
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

