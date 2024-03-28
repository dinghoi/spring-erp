<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)

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
dim sum_disabled_pay(6)
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
dim sum_hyubjo_amt(6)
dim sum_school_amt(6)
dim sum_nhis_bla_amt(6)
dim sum_long_bla_amt(6)
dim sum_deduct_tot(6)

view_condi=Request("view_condi")
pmg_yymm=request("pmg_yymm")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

savefilename = pmg_yymm + "월 상주회사별 급여현황.xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename


if view_condi = "" then
	view_condi = "케이원정보통신"
	curr_dd = cstr(datepart("d",now))
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	
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
	  com_tab(5) = "코리아디엔씨"
	  com_tab(6) = "합계"
      where_sql = " WHERE (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1')" 
   else  
      com_tab(1) = view_condi
	  com_tab(6) = "합계"
      where_sql = " WHERE (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"')"
end if   

sql = "select * from pay_month_give " + where_sql + order_sql

'Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    emp_no = rs("pmg_emp_no")
	pmg_company = rs("pmg_company")
	
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
                 sum_wetax(6) = sum_wetax(5) + de_wetax
	             sum_year_incom_tax(6) = sum_year_incom_tax(6) + de_year_incom_tax
                 sum_year_wetax(6) = sum_year_wetax(6) + de_year_wetax
				 sum_year_incom_tax2(6) = sum_year_incom_tax2(6) + de_year_incom_tax2
                 sum_year_wetax2(6) = sum_year_wetax2(6) + de_year_wetax2
                 sum_other_amt1(6) = sum_other_amt1(6) + de_other_amt1
                 sum_sawo_amt(6) = sum_sawo_amt(6) + de_sawo_amt
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
title_line = cstr(curr_yyyy) + "년 " + cstr(curr_mm) + "월 " + " 상주회사별 급여현황(" + view_condi + ")"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
				               <th rowspan="2" class="first" scope="col" >부문/상주회사</th>
                               <th rowspan="2" scope="col" >인원</th>
				               <th colspan="5" scope="col" style="background:#FFFFE6;">기본급여 및 제수당</th>
                               <th colspan="4" scope="col" style="background:#E0FFFF;">공제 및 차인지급액</th>
                               <th rowspan="2" scope="col" >지급액</th>
                               <th rowspan="2" scope="col" >비고</th>
			                </tr>
                            <tr>
								<td scope="col" style="background:#f8f8f8">기본급</td>
								<td scope="col" style="background:#f8f8f8;">식대</td>  
								<td scope="col" style="background:#f8f8f8;">연장근로수당</td>
                                <td scope="col" style="background:#f8f8f8;">통신비 등</td>
                                <td scope="col" style="background:#f8f8f8;">지급소계</td>
								<td scope="col" style="background:#f8f8f8;">4대보험</td>
                                <td scope="col" style="background:#f8f8f8;">소득세 등</td>
								<td scope="col" style="background:#f8f8f8;">기타공제등</td>
                                <td scope="col" style="background:#f8f8f8;">예수금계</td>
							</tr>
						</thead>
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
                                <td align="first"><%=rs("saup_count")%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("pmg_base_pay"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("pmg_meals_pay"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("pmg_overtime_pay"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(sub_give_hap,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("pmg_give_total"),0)%>&nbsp;</td>
                       <%  
                                  pmg_curr_pay = cdbl(rs("pmg_give_total")) - cdbl(rs("de_deduct_total"))
							  
							      hap_de_insur = cdbl(rs("de_nps_amt")) + cdbl(rs("de_nhis_amt")) + cdbl(rs("de_epi_amt")) + cdbl(rs("de_longcare_amt"))
							      hap_de_tax = cdbl(rs("de_income_tax")) + cdbl(rs("de_wetax")) + cdbl(rs("de_year_incom_tax")) + cdbl(rs("de_year_wetax")) + cdbl(rs("de_year_incom_tax2")) + cdbl(rs("de_year_wetax2"))
							      hap_de_other = cdbl(rs("de_other_amt1")) + cdbl(rs("de_sawo_amt")) + cdbl(rs("de_hyubjo_amt")) + cdbl(rs("de_school_amt")) + cdbl(rs("de_nhis_bla_amt")) + cdbl(rs("de_long_bla_amt"))
								  hap_deduct_tot = hap_de_insur + hap_de_tax + hap_de_other
                       %>
                                <td align="right"><%=formatnumber(hap_de_insur,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(hap_de_tax,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(hap_de_other,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(hap_deduct_tot,0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(pmg_curr_pay,0)%>&nbsp;</td>
                                <td align="right">&nbsp;</td>
							</tr>
					   <%
							rs.movenext()
						loop
						rs.close()
						
						sum_curr_pay = sum_give_tot(6) - sum_deduct_tot(6)
						
						sum_give_hap = sum_postage_pay(6) + sum_re_pay(6) + sum_car_pay(6) + sum_position_pay(6) + sum_custom_pay(6) + sum_job_pay(6) + sum_job_support(6) + sum_jisa_pay(6) + sum_long_pay(6) + sum_disabled_pay(6)
						sum_de_insur =sum_nps_amt(6) + sum_nhis_amt(6) + sum_epi_amt(6) + sum_longcare_amt(6)
						sum_de_tax =sum_income_tax(6) + sum_wetax(6) + sum_year_incom_tax(6) + sum_year_wetax(6) + sum_year_incom_tax2(6) + sum_year_wetax2(6)
						sum_de_other =sum_other_amt1(6) + sum_sawo_amt(6) + sum_hyubjo_amt(6) + sum_school_amt(6) + sum_nhis_bla_amt(6) + sum_long_bla_amt(6)
						
						%>
                          	<tr>
                                <th bgcolor="#EEFFFF" class="first">총계</th>
                                <th bgcolor="#EEFFFF" align="right"><%=formatnumber(pay_count(6),0)%>&nbsp;명</th>
                                <th bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_base_pay(6),0)%>&nbsp;</th>
                                <th bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_meals_pay(6),0)%>&nbsp;</th>
                                <th bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_overtime_pay(6),0)%>&nbsp;</th>
                                <th bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_give_hap,0)%>&nbsp;</th>
                                <th bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_give_tot(6),0)%>&nbsp;</th>
                                <th bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_de_insur,0)%>&nbsp;</th>
                                <th bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_de_tax,0)%>&nbsp;</th>
                                <th bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_de_other,0)%>&nbsp;</th>
                                <th bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_deduct_tot(6),0)%>&nbsp;</th>
                                <th bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_curr_pay,0)%>&nbsp;</th>
                                <th bgcolor="#EEFFFF" class="right">&nbsp;</th>
							</tr>
                         <%
						    for i = 1 to 6 
                        	     if	com_tab(i) <> "" then
								 
								    sum_curr_pay = sum_give_tot(i) - sum_deduct_tot(i)
						
						            sum_give_hap = sum_postage_pay(i) + sum_re_pay(i) + sum_car_pay(i) + sum_position_pay(i) + sum_custom_pay(i) + sum_job_pay(i) + sum_job_support(i) + sum_jisa_pay(i) + sum_long_pay(i) + sum_disabled_pay(i)
						            sum_de_insur =sum_nps_amt(i) + sum_nhis_amt(i) + sum_epi_amt(i) + sum_longcare_amt(i)
						            sum_de_tax =sum_income_tax(i) + sum_wetax(i) + sum_year_incom_tax(i) + sum_year_wetax(i) + sum_year_incom_tax2(i) + sum_year_wetax2(i)
						            sum_de_other =sum_other_amt1(i) + sum_sawo_amt(i) + sum_hyubjo_amt(i) + sum_school_amt(i) + sum_nhis_bla_amt(i) + sum_long_bla_amt(i)
						 %>	
                            <tr>
                                <td bgcolor="#FFE8E8" class="first"><%=com_tab(i)%></td>
                                <td bgcolor="#FFE8E8" align="right"><%=formatnumber(pay_count(i),0)%>&nbsp;명</td>
                                <td bgcolor="#FFE8E8" align="right"><%=formatnumber(sum_base_pay(i),0)%>&nbsp;</td>
                                <td bgcolor="#FFE8E8" align="right"><%=formatnumber(sum_meals_pay(i),0)%>&nbsp;</td>
                                <td bgcolor="#FFE8E8" align="right"><%=formatnumber(sum_overtime_pay(i),0)%>&nbsp;</td>
                                <td bgcolor="#FFE8E8" align="right"><%=formatnumber(sum_give_hap,0)%>&nbsp;</td>
                                <td bgcolor="#FFE8E8" align="right"><%=formatnumber(sum_give_tot(i),0)%>&nbsp;</td>
                                <td bgcolor="#FFE8E8" align="right"><%=formatnumber(sum_de_insur,0)%>&nbsp;</td>
                                <td bgcolor="#FFE8E8" align="right"><%=formatnumber(sum_de_tax,0)%>&nbsp;</td>
                                <td bgcolor="#FFE8E8" align="right"><%=formatnumber(sum_de_other,0)%>&nbsp;</td>
                                <td bgcolor="#FFE8E8" align="right"><%=formatnumber(sum_deduct_tot(i),0)%>&nbsp;</td>
                                <td bgcolor="#FFE8E8" align="right"><%=formatnumber(sum_curr_pay,0)%>&nbsp;</td>
                                <td bgcolor="#FFE8E8" class="right">&nbsp;</td>
							</tr>
                         <%
							     end if
						    next
					     %>
						</tbody>
					</table>
           </div>
		</div>				
	 </div>        				
  </body>
</html>

