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
dim sum_johab_amt(6)
dim sum_hyubjo_amt(6)
dim sum_school_amt(6)
dim sum_nhis_bla_amt(6)
dim sum_long_bla_amt(6)
dim sum_deduct_tot(6)

view_condi=Request("view_condi")
from_yymm=request("from_yymm")
to_yymm=request("to_yymm")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)
	
f_yymm = from_yymm
t_yymm = to_yymm

title_line = cstr(f_yymm) + " ∼ " + cstr(t_yymm) + "월 " + " 개인별 급여현황(기간별)-" + view_condi 

savefilename = cstr(f_yymm) + " ∼ " + cstr(t_yymm) + "월 개인별 급여현황(기간별).xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

	
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
		sum_johab_amt(i) = 0
        sum_hyubjo_amt(i) = 0
        sum_school_amt(i) = 0
        sum_nhis_bla_amt(i) = 0
        sum_long_bla_amt(i) = 0
        sum_deduct_tot(i) = 0
    next
	
	sum_curr_pay = 0	
	
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

Sql = "select * from pay_person_sum ORDER BY ps_company,ps_emp_no ASC" 
Rs.Open Sql, Dbconn, 1

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
				               <th rowspan="2" scope="col">성명</th>
                               <th rowspan="2" scope="col">사번</th>
                               <th rowspan="2" scope="col">직급</th>
                               <th rowspan="2" scope="col">회사</th>
                               <th rowspan="2" scope="col">부서</th>
				               <th colspan="6" scope="col" style="background:#FFFFE6;">지급항목</th>
                               <th colspan="7" scope="col" style="background:#FFFFE6;">공제항목</th>
                               <th rowspan="2" scope="col">지급액</th>
                               <th rowspan="2" scope="col">비고</th>
			                </tr>
                            <tr>
								<th scope="col">기본급</th>
								<th scope="col">식대</th>
								<th scope="col">통신비</th>
								<th scope="col">연장</th>
								<th scope="col">소급 등</th>
								<th scope="col">지급액계</th>
								<th scope="col">국민연금</th>
								<th scope="col">건강보험</th>
								<th scope="col">고용보험</th>
                                <th scope="col">장기요양</th>
                                <th scope="col">사우회비 등</th>
                                <th scope="col">소득세 등</th>
                                <th scope="col">공제액계</th>
							</tr>
						</thead>
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
                                <td align="center"><%=rs("ps_emp_name")%></td>
                                <td align="center"><%=rs("ps_emp_no")%></td>
                                <td align="center"><%=rs("ps_grade")%></td>
                                <td align="center"><%=rs("ps_company")%></td>
                                <td align="center"><%=rs("ps_org_name")%></td>
                                <td align="right"><%=formatnumber(rs("ps_base_pay"),0)%></td>
                                <td align="right"><%=formatnumber(rs("ps_meals_pay"),0)%></td>
                                <td align="right"><%=formatnumber(rs("ps_postage_pay"),0)%></td>
                                <td align="right"><%=formatnumber(rs("ps_overtime_pay"),0)%></td>
                                <td align="right"><%=formatnumber(hap_give_hap,0)%></td>
                                <td align="right"><%=formatnumber(rs("ps_give_total"),0)%></td>
                                
                                <td align="right"><%=formatnumber(rs("ps_nps_amt"),0)%></td>
                                <td align="right"><%=formatnumber(rs("ps_nhis_amt"),0)%></td>
                                <td align="right"><%=formatnumber(rs("ps_epi_amt"),0)%></td>
                                <td align="right"><%=formatnumber(rs("ps_longcare_amt"),0)%></td>
                                <td align="right"><%=formatnumber(hap_de_other,0)%></td>
                                <td align="right"><%=formatnumber(hap_de_tax,0)%></td>
                                <td align="right"><%=formatnumber(rs("ps_deduct_total"),0)%></td>

                                <td align="right"><%=formatnumber(pmg_curr_pay,0)%></td>
                                <td align="right">&nbsp;</td>
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
                                <td colspan="4" align="center" style="background:#ffe8e8;">총계</td>
                                <td align="center" style="background:#ffe8e8;"><%=formatnumber(pay_count(6),0)%>&nbsp;명</td>
                                <td align="right" style="background:#ffe8e8;"><%=formatnumber(sum_base_pay(6),0)%></td>
                                <td align="right" style="background:#ffe8e8;"><%=formatnumber(sum_meals_pay(6),0)%></td>
                                <td align="right" style="background:#ffe8e8;"><%=formatnumber(sum_postage_pay(6),0)%></td>
                                <td align="right" style="background:#ffe8e8;"><%=formatnumber(sum_overtime_pay(6),0)%></td>
                                <td align="right" style="background:#ffe8e8;"><%=formatnumber(hap_give_hap,0)%></td>
                                <td align="right" style="background:#ffe8e8;"><%=formatnumber(sum_give_tot(6),0)%></td>
                                
                                <td align="right" style="background:#ffe8e8;"><%=formatnumber(sum_nps_amt(6),0)%></td>
                                <td align="right" style="background:#ffe8e8;"><%=formatnumber(sum_nhis_amt(6),0)%></td>
                                <td align="right" style="background:#ffe8e8;"><%=formatnumber(sum_epi_amt(6),0)%></td>
                                <td align="right" style="background:#ffe8e8;"><%=formatnumber(sum_longcare_amt(6),0)%></td>
                                <td align="right" style="background:#ffe8e8;"><%=formatnumber(hap_de_other,0)%></td>
                                <td align="right" style="background:#ffe8e8;"><%=formatnumber(hap_de_tax,0)%></td>
                                <td align="right" style="background:#ffe8e8;"><%=formatnumber(sum_deduct_tot(6),0)%></td>
                                
                                <td align="right" style="background:#ffe8e8;"><%=formatnumber(sum_curr_pay,0)%></td>
                                <td align="right" style="background:#ffe8e8;">&nbsp;</td>
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
                                <td colspan="4" align="center" style="background:#eeffff;"><%=com_tab(i)%></td>
                                <td align="center" style="background:#eeffff;"><%=formatnumber(pay_count(i),0)%>&nbsp;명</td>
                                <td align="right" style="background:#eeffff;"><%=formatnumber(sum_base_pay(i),0)%></td>
                                <td align="right" style="background:#eeffff;"><%=formatnumber(sum_meals_pay(i),0)%></td>
                                <td align="right" style="background:#eeffff;"><%=formatnumber(sum_postage_pay(i),0)%></td>
                                <td align="right" style="background:#eeffff;"><%=formatnumber(sum_overtime_pay(i),0)%></td>
                                <td align="right" style="background:#eeffff;"><%=formatnumber(hap_give_hap,0)%></td>
                                <td align="right" style="background:#eeffff;"><%=formatnumber(sum_give_tot(i),0)%></td>
                                
                                <td align="right" style="background:#eeffff;"><%=formatnumber(sum_nps_amt(i),0)%></td>
                                <td align="right" style="background:#eeffff;"><%=formatnumber(sum_nhis_amt(i),0)%></td>
                                <td align="right" style="background:#eeffff;"><%=formatnumber(sum_epi_amt(i),0)%></td>
                                <td align="right" style="background:#eeffff;"><%=formatnumber(sum_longcare_amt(i),0)%></td>
                                <td align="right" style="background:#eeffff;"><%=formatnumber(hap_de_other,0)%></td>
                                <td align="right" style="background:#eeffff;"><%=formatnumber(hap_de_tax,0)%></td>
                                <td align="right" style="background:#eeffff;"><%=formatnumber(sum_deduct_tot(i),0)%></td>
                                
                                <td align="right" style="background:#eeffff;"><%=formatnumber(sum_curr_pay,0)%></td>
                                <td align="right" style="background:#eeffff;">&nbsp;</td>
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

